using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace Ryr.ExcelExport
{

    [Export(typeof(IXrmToolBoxPlugin))]
    [ExportMetadata("BackgroundColor", "White")]
    [ExportMetadata("BigImageBase64", "iVBORw0KGgoAAAANSUhEUgAAAFAAAABQCAYAAACOEfKtAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAFiUAABYlAUlSJPAAAAACYktHRAD/h4/MvwAAAAl2cEFnAAAAgAAAAIAAMOExmgAAACV0RVh0ZGF0ZTpjcmVhdGUAMjAxMC0wMi0xMVQxMjo1MDoxNy0wNjowMFE4eUIAAAAldEVYdGRhdGU6bW9kaWZ5ADIwMDktMTAtMjJUMjM6MjA6MDAtMDU6MDDtiMMIAAANGklEQVR4Xu2cBYzUTBvHB3d3dzuCu2twCO7BJbhDgrtLcII7AY7gBAvuFtzd3d32u99D576+9x5wK91dXvafNNtOp73pv4/P9ELZAqB8cBihjV8fHISPQCfhI9BJWGYD161bp6ZOnaq+f/+uQoUKZbS6F+HChVMxYsRQOXLkUNWrV1dp0qQxzrgOlhB49OhRlTdvXhUmTBgVPnx4jxHIo/ECP336JMeRI0dWY8eOVU2aNFGRIkWSNmdhCYGTJk1SnTp1UqlSpVLRokUTAi34MyEGf/vLly/q7du36vHjxyplypRq1qxZqmjRokYPx2EJgdeuXVPp06dXadOmFSn0FHhxoUOHVmHDhg0cx7dv39SrV6/UrVu31MiRI1WPHj2k3VFYZgNPnz6tli9fLm/eEyoMca9fv1aPHj1S9+/fV8+ePRO1hUjG8/XrV3Xv3j3VqFEjNW7cOOMq+2EZgd4EiMKpTZ48WdQZIrVZuXHjhlq4cKGqXLmy0ds+/BUEmoHE4eTMJKLSmzdvVunSpTN6hRx/XRyItJUrV04cCuRBIuHOqFGjjB72wWkJPHbsmAzqzJkz6sOHDx6xd8GBkKVgwYKqfv36KlOmTEbr/1GpUiV18+ZNFSFChMBQZ+3atXZLoVMEMghEH8/GQHiT3gJIwYHhTDp06KDGjBljnPmBEydOqIYNG6qIESPK8fv371WLFi1U9+7d5TikcJhAyNu4caOKGjWqihUrlsR7EKjtiqfBOLRkIWlt27ZV48ePN87+QKlSpSQuJNinH5K6cuVK42zI4BCBnTt3VhMnTlTx48eXDeljAIQG3kIesZ85C8ITI4VIncacOXPUwIED5RkgGykk/OJ5Qgq7CeSPJEqUSIhKkSKFSB3k5cmTR2XIkEHaPUmilryTJ0+qc+fOBXpb1BltOXDggNFTqePHj4sUEvBzzdOnT9WRI0dUwoQJjR4hAATaA39/f1uAXbEFpEO2XLly2fz8/Gzz5s0zznoXArIMW4Ba2nLkyGHLmTOnLYAY28uXL42zNtvly5dtMWPGlHNsyZMnt128eNE4GzLYHcaQAmGYsXk4jyhRokhy7o1o2bKlSGDAc8oxqkmaaYazZsduAiEN24KN4Q9rG+ON4EXreA8wVsZvBqrL5igcCqQhj8F5OyAGApEyDfMLZ98Z6QMuZYGB6jf6uw1JMEuDPg6urzNq9juSnNUglxF4/fp18XLYxHjx4qnEiRMHu+HhsJ9IMZkCwEMWKlRI2szX0peKMp5+/vz50tfb4DICCWl2794tkTwPHuDt1OfPn4UcvRHuUE7q2rWr2rNnj1q6dKlcC0GkgwsWLFCpU6eWUIm+oEGDBmrfvn1e66jsjgMJRocOHSo5I6qFZzt8+LBx9gcgi7Ro//79IpUa9KdKvX79eqPl36DkNGXKFEmx8ufPr+rVq6fevXsneTb2jJfyK/A4SHLdunWlFkjVGSnmfs+fP5dMI3fu3NL36tWrKlu2bPIsvFjiwK1bt0o8G1JYQqAG0kPpCLUG/ClI6Natm2rTpo20mUFBonjx4vLAsWPHFrIoir548UJIDOpBfwaihLNnz8rYMBNWEmipK6VkHj16dCEaYLCpkvASkNKgqFatmqg+gGxUG/JJtZBcMgYelu1n+/TjF/LslA2HYCmByZIlU0OGDFFPnjwJfBjeNCrWrl07OdZo3LixBLm6KIH0YQfxwgTDbLQjSWzs0wZRQdvZgLMeNiSwlEBQvnx5VaVKFbFhmkSkEFVZsWKFHJPQ40Sohly5ckVdunRJsggk8sKFC6pixYpSTTl//rzcr3379rJPGR5TwH6tWrWk2sy1mI2AtCxYKXc1LCcQMBeLHTKrMpI2YcIEtWjRIjVo0CCRpDhx4oiEAvrq/hBh3te2kH0kFNCm++gg3x0qbKkTMePUqVOqatWqYhP1A/LQDx8+VB8/fhQ7B8mQ++DBA3EcHMeNG1eMPeeRXOJNqkHso/IJEiSQe9KHvqgxL4Ix0aadEu1/nBMxg4FSXifGM9vDJEmSSNCMs4CALVu2SO2ub9++0h/P3KxZM6lBos5NmzZVXbp0EbVt3bq1tONxqTpzDCnEjcAdEug2AsGwYcMkUNYqyAMijcSK/BKyoO59+vSRQBvy8OT87t27Vw0fPlzIImBHC6j50c59qe0dOnRI9evXL3Ce9z/hRIKCZRVIoYbOdzWZpHJs2qahrnhjyCDDQd2RXFQXT8015j60o7LuQljj1y2YMWOGmjlzpqyWgghzrMYvtow0D2DjsI86vEFySeewgSwbweOy4gBvi1oTKmEG8NAa7lBhtzkRlnmQXvGApGuFCxcWo65JRHq0V2Wim8wDQlBviIwZM6bEd6RnOAliSSaEmNAi89B9kEgk9T/lRLBnkJcvXz4hD2DrIA3C+NXvMWPGjOJESPdq1qwpcV2dOnWEeB6Y+3Ts2FH2SRVpJ3YkbsTZUDHfuXOn3MsdEmg5gbz1YsWKiV3D+GtQKMDLIm0AEpEqPC3xIdexIIigmPQOUnES7N+9e1cmhwh12KcdaUQ6d+zYoQ4ePBh4T6thqQ0kniOZh5jt27eLupqBtyVsgSg9BUkfbCUqmTRpUiGTqg5qic0jRCH0uXPnjrTjQOjLPgUI+mBLmW1zByyVQFIxJrWnTZumsmbNarT+E6zPw/ZpdYNEHAj2jLiOmiGlMVSa1I99rsGWkd7Rjn3F+ZBfswpr8eLFci93qLAlTgSJKlmypExSZ86cWWK3X4EMBZumiwAA24jxx4kQB0IqDgMVx4kQdGMfkTpCl8uXL0tffYxk/3FOhAFCMGEK5DFwVBOCCTPMIIZD3bZt2yZE8DJ4lzouBNg4CKOAgPOZPn262rVrl1qzZo2EQzgLFgSh8txnw4YNau7cueKggDsk0GUEIiVIRs+ePcU+YYc4JhzBHpHLUpHR6NWrl9iyMmXKqIsXL4ozIRgmPtQbJCOVmzZtkiVpkMzLwRxkz55d5cyZM3A/V65csk9blixZjL9iPVxGIIOGIEhgQ42RLNQXSeIc9TsNJBWSdP/bt2+L1NLfvLE8AzvK9Vr1vAkuVWHsDCrLhudl0wVOzpnBOeyn7q/7Bbfpvu4IS+yFpV74b4CPQCfhI9BJ+Ah0Ej4CnYSPQCfhI9BJWEIgKRR1OaokFAo0yCQoORE0g2XLlknB4XcfuXAfNlJFM8hdqfgETRPdCUsIJOCtXbu2lJpI1TSonlCWYkE6oNZHXkup61cg5+VelPWZSAKU+akxkg4yseQpWKbCVI3JHkj+dXGA5WuAmTagswyWgPwKVKGZjHrz5o1UU4D+5oPcl6/RPQXLCKTczqJJHpwPsGfPni3FgbJly8okEEDVdbpnBlJGVVmrOhJNYZWFQ9QHWeG1evVqOUf1xZOw1IlQ+KRQgApSVAV61g3ocpP+hRhKV/y7ACrZLNrUJX9qhtT6IJt5EOwefX9WqHUXLCWQElSBAgWkKoPBL1269D9sYlCwGoFv77CR/fv3l1/zHC9TAEgxcyBUrAcMGGCc8RwsJRAUKVJEHppqDKu0fgW+/ERdIZoFR3w1ZAZToXr5G9LJLJ+nYSmBhC3YKIqqlMz1muifoUKFCmITlyxZIktz+ZbNDNpZE0MdEYfi7+9vnPEcLCWQtc5UqpnWpOBKpXr06NHG2X+jd+/eMiHE/3ehiMoEEl4cIMVUsfHoLCjCbo4YMUJekidhKYE6XGHlVPPmzeWhmTXDHgYHyCF+ZEKIaQCmNbF3QFewUVtm31jRRVDO+kJHwXi0A3MUlhGIpPGAeE7sIOELcyRIFv/RCEAYEqTjRLISZtFYYcVqK1Rfl/EHDx4sv/qDaFatci0TStpT2wvsLX9bk0jEEDSk+h0sIRB1mzdvnmQKqKVGq1atJBNZtWqVHONhOWYCCpCqQRySS4qGNOJQ+DaZBytRokSgF2dpL9kJJDIv7AgwFQThhET8bewu2Y49sGRemHaWXOB5IREHotuZYGLpGpLJGyf2Y85Xhyusc8FBcJ3OUMiDGSaTUoQvGtxL59rM+gUF94L0n80LA8bEC0UakWr9SUZI4ZAE/o5zPCkZCNKlyQO0QwztTCQRkiBF5liPF0N6Zk7v6E8/M3mAiXaIC468kIIxsXgJabeXPGA3gXYKrEfBUjckywxXj99uAlFZBsZAGBwqgLH3RmDbWM3FeAEOwzw37QrYbQNZ61ejRo1AFcNhkCHgJVmb4g0SyovFtpJP44wwC5AIocSiroTdBAK8F4YfO8blrHnGmEOgt4ClwIxL1wzZp/DgqMf+GRwikAIoWQJGnLfNRpCLl+PXG4Dz0svkGB8SSPGWHNqVcIhAQDDMKiikjhiNQQIdFHsDGBMmho2aJNUhV8NhAgHLych3WQSEcdbG2htArAlxlMT44NHPz88441o4RaAGa/hYYksArCXRU9CPQ6Cu82kr4RIC/2Z4j879ofAR6CR8BDoJH4FOwkegU1Dqf6RfPxF+6r53AAAAAElFTkSuQmCC")]
    [ExportMetadata("Description", "Exports all the records from the selected view to Excel")]
    [ExportMetadata("Name", "Export To Excel")]
    [ExportMetadata("PrimaryFontColor", "Black")]
    [ExportMetadata("SecondaryFontColor", "Gray")]
    [ExportMetadata("SmallImageBase64", "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")]
    public class Plugin : PluginBase
    {
        public Plugin()
        {
            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(AssemblyResolveEventHandler);
        }

        public override IXrmToolBoxPluginControl GetControl()
        {
            return new ExcelExportPlugin();
        }

        //https://github.com/albanian-xrm/Early-Bound/blob/c1b9617f681410e5ea508e32626fca4b2c9371a5/AlbanianXrm.EarlyBound/MyPlugin.cs#L39-L85
        private Assembly AssemblyResolveEventHandler(object sender, ResolveEventArgs args)
        {
            Assembly loadAssembly = null;
            Assembly currAssembly = Assembly.GetExecutingAssembly();

            // base name of the assembly that failed to resolve
            var argName = args.Name.Substring(0, args.Name.IndexOf(",", StringComparison.InvariantCulture));

            // check to see if the failing assembly is one that we reference.
            var refAssemblies = currAssembly.GetReferencedAssemblies().ToList();
            var refAssembly = refAssemblies.FirstOrDefault(a => a.Name == argName);

            // if the current unresolved assembly is referenced by our plugin, attempt to load
            if (refAssembly == null) return loadAssembly;

            // load from the path to this plugin assembly, not host executable
            string dir = Path.GetDirectoryName(currAssembly.Location).ToUpperInvariant();
            string folder = Path.GetFileNameWithoutExtension(currAssembly.Location);
            dir = Path.Combine(dir, folder);

            var assemblyPath = Path.Combine(dir, $"{argName}.dll");

            loadAssembly = File.Exists(assemblyPath)
                ? Assembly.LoadFrom(assemblyPath)
                : throw new FileNotFoundException($"Unable to locate dependency: {assemblyPath}");

            return loadAssembly;
        }
    }
}
