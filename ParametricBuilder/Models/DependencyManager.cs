using System.Collections.Generic;
using System.Linq;

namespace ParametricBuilder.Models
{
    public static class DependencyManager
    {
        private static readonly Dictionary<string, List<ParameterModel>> Dependencies =
            new Dictionary<string, List<ParameterModel>>();

        public static void RegisterDependency(string controllerName, ParameterModel dependent)
        {
            if (!Dependencies.ContainsKey(controllerName))
            {
                Dependencies[controllerName] = new List<ParameterModel>();
            }

            if (!Dependencies[controllerName].Contains(dependent))
            {
                Dependencies[controllerName].Add(dependent);
                dependent.Controller.HasDependents = true;
            }
        }

        public static void NotifyDependents(string controllerName)
        {
            if (Dependencies.TryGetValue(controllerName, out var dependents))
            {
                foreach (var dependent in dependents.ToList()) // ToList for safe iteration
                {
                    dependent.UpdateState();
                }
            }
        }
    }
}