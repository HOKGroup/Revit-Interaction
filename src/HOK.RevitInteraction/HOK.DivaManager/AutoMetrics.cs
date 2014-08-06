using System;
using Rhino;
using Rhino.Commands;

namespace HOK.DivaManager
{
    [System.Runtime.InteropServices.Guid("02033515-a9cc-4be0-9bc3-59592308e57b")]
    public class AutoMetrics : Command
    {
        static AutoMetrics _instance;
        public AutoMetrics()
        {
            _instance = this;
        }

        ///<summary>The only instance of the AutoMetrics command.</summary>
        public static AutoMetrics Instance
        {
            get { return _instance; }
        }

        public override string EnglishName
        {
            get { return "AutoMetrics"; }
        }

        protected override Result RunCommand(RhinoDoc doc, RunMode mode)
        {
            // TODO: complete command.
            return Result.Success;
        }
    }
}
