using System.Buffers.Binary;
using System.Text;
using Gongwen.Desktop;

int passed=0,failed=0;
void Check(bool condition,string reason){if(!condition)throw new Exception(reason);}
async Task Test(string name,Func<Task> action){try{await action();passed++;Console.WriteLine("PASS "+name);}catch(Exception ex){failed++;Console.WriteLine("FAIL "+name+": "+ex.GetType().Name+" "+ex.Message);}}
Task Sync(Action action){action();return Task.CompletedTask;}
async Task Reject(Func<Task> action){try{await action();}catch(Exception ex)when(ex is InvalidDataException or EndOfStreamException){return;}throw new Exception("Expected bounded input rejection");}
await Test("CLI_default_and_open_contract",()=>Sync(()=>{
 Check(OfficeImportCommand.TryParse([],out var normal,out _)&&normal.Mode==OfficeImportLaunchMode.Normal,"Normal startup rejected");
 Check(OfficeImportCommand.TryParse(["--open",@"c:\文档 资料\示例.docx"],out var open,out _)&&open.Mode==OfficeImportLaunchMode.Open&&open.DocumentPath==@"C:\文档 资料\示例.docx","Open startup not normalized");
 Check(OfficeImportCommand.TryParse(["--smoke-test","sample.docx","output.docx","result.json"],out var smoke,out _)&&smoke.Mode==OfficeImportLaunchMode.SmokeTest,"Original smoke mode broke");
}));
await Test("CLI_unknown_and_extra_arguments_rejected",()=>Sync(()=>{
 foreach(var args in new[]{new[]{"--open"},new[]{"--open",@"C:\a.docx","--smoke-test"},new[]{"--run","cmd.exe"},new[]{"--smoke-test","a","b"}})
  Check(!OfficeImportCommand.TryParse(args,out _,out _),"Unexpected CLI accepted");
}));
await Test("UNC_and_dot_segments_normalized",()=>Sync(()=>{
 Check(OfficeImportCommand.TryValidateDocumentPath(@"\\fileserver\共享\资料\..\示例.docx",out var unc)&&unc==@"\\fileserver\共享\示例.docx","UNC changed");
 Check(OfficeImportCommand.TryValidateDocumentPath("c:/docs/./sample.DOCX",out var path)&&path==@"C:\docs\sample.DOCX","Slash normalization failed");
}));
await Test("URI_device_ADS_relative_and_wrong_extension_rejected",()=>Sync(()=>{
 foreach(var input in new[]{"https://example.invalid/a.docx","file:///C:/a.docx",@"C:a.docx",@"\a.docx",@"..\a.docx",@"C:\a.docx:stream",@"\\?\C:\a.docx",@"\\.\C:\a.docx",@"\\host\share",@"C:\a.docm",@"C:\a.doc",@"C:\CON.docx",@"C:\LPT1.docx",@"C:\..\a.docx",@"C:\bad.\a.docx",@"C:\bad \a.docx"})
  Check(!OfficeImportCommand.TryValidateDocumentPath(input,out _),"Unsafe path accepted");
}));
await Test("control_invalid_UTF16_and_path_length_rejected",()=>Sync(()=>{
 foreach(var input in new[]{"C:\\a\n.docx","C:\\a\0.docx","C:\\a\u2028.docx","C:\\a\ud800.docx","C:\\"+new string('a',32760)+".docx"})
  Check(!OfficeImportCommand.TryValidateDocumentPath(input,out _),"Unsafe argument encoding accepted");
}));
await Test("framed_path_round_trip",async()=>{
 using var memory=new MemoryStream();await OfficeImportProtocol.WritePathAsync(memory,@"C:\虚构 资料\a & b.docx",default);memory.Position=0;
 Check(await OfficeImportProtocol.ReadPathAsync(memory,default)==@"C:\虚构 资料\a & b.docx","Path frame changed");
});
await Test("oversized_negative_and_truncated_frames_rejected",async()=>{
 foreach(var length in new[]{0,-1,int.MaxValue,OfficeImportProtocol.MaxPayloadBytes+1}){
  var bytes=new byte[4];BinaryPrimitives.WriteInt32LittleEndian(bytes,length);using var stream=new MemoryStream(bytes);await Reject(async()=>{_ = await OfficeImportProtocol.ReadPathAsync(stream,default);});
 }
 using var shortStream=new MemoryStream(new byte[]{1,0});await Reject(async()=>{_ = await OfficeImportProtocol.ReadPathAsync(shortStream,default);});
});
await Test("bad_UTF8_and_nonpath_payload_rejected",async()=>{
 foreach(var payload in new[]{new byte[]{0xc3,0x28},Encoding.UTF8.GetBytes("https://example.invalid/a.docx")}){
  using var stream=new MemoryStream();var size=new byte[4];BinaryPrimitives.WriteInt32LittleEndian(size,payload.Length);stream.Write(size);stream.Write(payload);stream.Position=0;
  await Reject(async()=>{_ = await OfficeImportProtocol.ReadPathAsync(stream,default);});
 }
});
await Test("delivery_status_round_trip_and_unknown_rejected",async()=>{
 foreach(var expected in new[]{OfficeImportDelivery.Accepted,OfficeImportDelivery.Busy,OfficeImportDelivery.Rejected}){
  using var stream=new MemoryStream();await OfficeImportProtocol.WriteDeliveryAsync(stream,expected,default);stream.Position=0;
  Check(await OfficeImportProtocol.ReadDeliveryAsync(stream,default)==expected,"Status changed");
 }
 using var invalid=new MemoryStream(new byte[]{99});await Reject(async()=>{_ = await OfficeImportProtocol.ReadDeliveryAsync(invalid,default);});
});
await Test("pipe_names_separate_user_and_session_without_identity_text",()=>Sync(()=>{
 var a=OfficeImportPipeName.ForIdentity("fictional-user-A",1);var b=OfficeImportPipeName.ForIdentity("fictional-user-B",1);var c=OfficeImportPipeName.ForIdentity("fictional-user-A",2);
 Check(a!=b&&a!=c&&!a.Contains("fictional-user",StringComparison.Ordinal),"Pipe identity separation failed");
}));
await Test("real_local_pipe_delivers_path_and_busy_then_survives_bad_client",async()=>{
 string name="gw"+Guid.NewGuid().ToString("N")[..12];string? delivered=null;
 using var server=new OfficeImportPipeServer(name,(path,_)=>{delivered=path;return Task.FromResult(path.EndsWith("busy.docx",StringComparison.Ordinal)?OfficeImportDelivery.Busy:OfficeImportDelivery.Accepted);});
 var loop=server.RunAsync();
 Check(await OfficeImportPipeClient.SendAsync(name,@"C:\fictional\sample.docx")==OfficeImportDelivery.Accepted&&delivered==@"C:\fictional\sample.docx","Actual local pipe did not deliver");
 Check(await OfficeImportPipeClient.SendAsync(name,@"C:\fictional\busy.docx")==OfficeImportDelivery.Busy,"Busy state lost");
 using(var bad=new System.IO.Pipes.NamedPipeClientStream(".",name,System.IO.Pipes.PipeDirection.InOut,System.IO.Pipes.PipeOptions.Asynchronous|System.IO.Pipes.PipeOptions.CurrentUserOnly)){
  using var time=new CancellationTokenSource(TimeSpan.FromSeconds(5));await bad.ConnectAsync(time.Token);await bad.WriteAsync(new byte[]{0xff,0xff,0xff,0x7f},time.Token);await bad.FlushAsync(time.Token);
  Check(await OfficeImportProtocol.ReadDeliveryAsync(bad,time.Token)==OfficeImportDelivery.Rejected,"Malformed client not rejected");
 }
 Check(await OfficeImportPipeClient.SendAsync(name,@"C:\fictional\again.docx")==OfficeImportDelivery.Accepted,"Bad client killed listener");
 server.Dispose();await loop.WaitAsync(TimeSpan.FromSeconds(3));
});
await Test("rapid_sequential_clients_keep_listener_and_ACKs_alive",async()=>{
 string name="gw"+Guid.NewGuid().ToString("N")[..12];int received=0;
 using var server=new OfficeImportPipeServer(name,(_,_)=>{Interlocked.Increment(ref received);return Task.FromResult(OfficeImportDelivery.Accepted);});
 var loop=server.RunAsync();
 for(int index=0;index<80;index++)
  Check(await OfficeImportPipeClient.SendAsync(name,@"C:\fictional\sequential-"+index+".docx")==OfficeImportDelivery.Accepted,"Sequential client lost its acknowledgement");
 Check(Volatile.Read(ref received)==80,"Sequential clients lost a request");
 server.Dispose();await loop.WaitAsync(TimeSpan.FromSeconds(3));
});
Console.WriteLine($"RESULT {passed} passed, {failed} failed; real local IPC exercised, WPS/Word host not exercised");
return failed==0?0:1;
