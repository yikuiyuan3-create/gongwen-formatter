using System;

namespace Gongwen.AddinRegistrar;

internal interface IUserRegistry
{
    void CreateOrUpdate(RegistrationView view, RegistrationNode node);
    void DeleteOwnedTree(RegistrationView view, string path);
}

internal static class RegistrationRunner
{
    internal static int Execute(RegistrationOperation operation, AddinRegistrationMetadata metadata, IUserRegistry registry)
    {
        RegistrationPlan plan;
        try { plan = RegistrationPlanBuilder.Build(metadata); }
        catch (Exception) { return 2; }
        var views = new[] { RegistrationView.Registry32, RegistrationView.Registry64 };
        if (operation == RegistrationOperation.Register)
        {
            try
            {
                foreach (var view in views)
                    foreach (var node in plan.Nodes) registry.CreateOrUpdate(view, node);
                return 0;
            }
            catch (Exception)
            {
                // Best effort in both views, even if the first failure happened in
                // the other view. Every path is a fixed product-owned subtree.
                DeleteOwned(plan, views, registry);
                return 2;
            }
        }
        return operation == RegistrationOperation.Unregister && DeleteOwned(plan, views, registry) ? 0 : 2;
    }

    private static bool DeleteOwned(RegistrationPlan plan, RegistrationView[] views, IUserRegistry registry)
    {
        var success = true;
        foreach (var view in views)
            foreach (var path in plan.UnregisterPaths)
                try { registry.DeleteOwnedTree(view, path); }
                catch (Exception) { success = false; }
        return success;
    }
}
