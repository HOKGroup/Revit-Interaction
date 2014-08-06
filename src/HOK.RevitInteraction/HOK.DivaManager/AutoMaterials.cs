using System;
using Rhino;
using Rhino.Commands;

namespace HOK.DivaManager
{
    [System.Runtime.InteropServices.Guid("d54d270f-1e5a-48bb-bcb8-e2fd4aa7f614")]
    public class AutoMaterials : Command
    {
        static AutoMaterials _instance;
        public AutoMaterials()
        {
            _instance = this;
        }

        ///<summary>The only instance of the AutoMaterials command.</summary>
        public static AutoMaterials Instance
        {
            get { return _instance; }
        }

        public override string EnglishName
        {
            get { return "AutoMaterials"; }
        }

        protected override Result RunCommand(RhinoDoc doc, RunMode mode)
        {
            // TODO: complete command.

            return Result.Success;
        }
    }
}
