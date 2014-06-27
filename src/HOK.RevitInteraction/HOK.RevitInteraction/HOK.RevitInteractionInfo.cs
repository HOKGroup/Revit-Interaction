using System;
using System.Drawing;
using Grasshopper.Kernel;

namespace HOK.RevitInteraction
{
  public class RevitInteractionInfo : GH_AssemblyInfo
  {
    public override string Name
    {
      get
      {
        return "HOKRevitInteraction";
      }
    }
    public override Bitmap Icon
    {
      get
      {
        //Return a 24x24 pixel bitmap to represent this GHA library.
        return null;
      }
    }
    public override string Description
    {
      get
      {
        //Return a short string describing the purpose of this GHA library.
        return "";
      }
    }
    public override Guid Id
    {
      get
      {
        return new Guid("2c648cff-7fbd-4e15-a48a-6c702971ee7f");
      }
    }

    public override string AuthorName
    {
      get
      {
        //Return a string identifying you or your company.
        return "HOK Group";
      }
    }
    public override string AuthorContact
    {
      get
      {
        //Return a string representing your preferred contact details.
        return "";
      }
    }
  }
}
