namespace Gongwen.OfficeAddin;

internal static class RibbonDefinition
{
    // Office 2007 schema is intentionally used for the broadest Ribbon host
    // compatibility. Icons, scripts, websites and remote resources are absent.
    internal const string Xml = """
        <customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
          <ribbon>
            <tabs>
              <tab id="GongwenFormattingTab" label="公文排版">
                <group id="GongwenFormattingGroup" label="公文一键排版">
                  <button id="GongwenFormatCurrent" label="排版当前文档" size="large" onAction="FormatCurrentDocument"
                          screentip="导入已保存的当前文档"
                          supertip="请先按 Ctrl+S 保存 DOCX。排版工具读取磁盘文件并另存新副本，原稿保持不变。" />
                  <button id="GongwenOpenTool" label="打开排版工具" onAction="OpenTool" />
                  <button id="GongwenHelp" label="使用说明" onAction="ShowHelp" />
                </group>
              </tab>
            </tabs>
          </ribbon>
        </customUI>
        """;
}
