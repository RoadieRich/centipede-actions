using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using CentipedeInterfaces;
using Microsoft.VisualBasic.FileIO;


namespace ShellActions
{
    public abstract class ShellAction : Centipede.Action
    {
        protected ShellAction(string name, IDictionary<string, object> v, ICentipedeCore core)
                : base(name, v, core)
        { }

        protected static DirectoryInfo Pwd = new DirectoryInfo(Environment.CurrentDirectory);
    }

    [ActionCategory("Shell", DisplayName = "Set Current Directory")]
    public class SetCwd : ShellAction
    {
        public SetCwd(IDictionary<string, object> v, ICentipedeCore core)
            : base("Set Current Directory", v, core)
        { }

        [ActionArgument]
        public String DirectoryPath = "";

        protected override void DoAction()
        {
            try
            {
                Pwd = new DirectoryInfo(ParseStringForVariable(DirectoryPath));
            }
            catch (DirectoryNotFoundException e)
            {
                throw new ActionException(string.Format("Directory not found: {0}", DirectoryPath), e, this);
            }
        }
    }

    [ActionCategory("Shell", DisplayName = "Change current directory")]
    public class Cd : ShellAction
    {
        public Cd(IDictionary<string, object> v, ICentipedeCore core)
            : base("Change current directory", v, core)
        { }

        [ActionArgument]
        public String DirectoryName = "";

        protected override void DoAction()
        {

            try
            {
                Pwd = Pwd.GetDirectories(ParseStringForVariable(DirectoryName)).First();
            }
            catch (DirectoryNotFoundException e)
            {
                throw new ActionException(string.Format("Directory not found: {0}", DirectoryName), e, this);
            }
        }
    }

    [ActionCategory("Shell", DisplayName="Create Directory")]
    public class MkDir : ShellAction
    {
        public MkDir(IDictionary<string, object> v, ICentipedeCore core)
                : base("Create Directory", v, core)
        { }

        [ActionArgument]
        public String DirectoryName = "";

        [ActionArgument]
        public Boolean NavigateInto;
        
        protected override void DoAction()
        {
            DirectoryInfo newDir = Pwd.CreateSubdirectory(ParseStringForVariable(DirectoryName));
            if (NavigateInto)
            {
                Pwd = newDir;
            }
        }
    }

    [ActionCategory("Shell", DisplayName = "Copy File")]
    public class Cp : ShellAction
    {
        public Cp(IDictionary<string, object> v, ICentipedeCore core)
                : base("Copy File", v, core)
        { }

        [ActionArgument]
        public String From = "";
        
        [ActionArgument]
        public String To = "";

        protected override void DoAction()
        {
            string from = ParseStringForVariable(this.From);

            DirectoryInfo directory;
            directory = Path.IsPathRooted(@from) ? FileSystem.GetDirectoryInfo(@from) : Pwd;
            foreach (FileInfo file in directory.EnumerateFiles(@from))
            {
                file.CopyTo(ParseStringForVariable(this.To), this.AllowOverWrite);
            }
        }

        [ActionArgument]
        public bool AllowOverWrite;

    }

    [ActionCategory("Shell", DisplayName = "Move File")]
    public class Mv : ShellAction
    {
        public Mv(IDictionary<string, object> v, ICentipedeCore core)
            : base("Move File", v, core)
        { }

        [ActionArgument]
        public String From = "";

        [ActionArgument]
        public String To = "";

        protected override void DoAction()
        {
            foreach (FileInfo file in Pwd.EnumerateFiles(ParseStringForVariable(From)))
            {
                file.MoveTo(ParseStringForVariable(To));
            }
        }
    }

    [ActionCategory("Shell", DisplayName = "Get Current Directory Name")]
    public class Pwd : ShellAction
    {
        public Pwd(IDictionary<string, object> v, ICentipedeCore core)
            : base("Get Current Directory Name", v, core)
        { }

        [ActionArgument(Literal=true)]
        public String DestinationVariable = "";

        protected override void DoAction()
        {
            Variables[DestinationVariable] = Pwd.FullName;
        }
    }

    [ActionCategory("Shell", DisplayName = "Delete File")]
    public class Del : ShellAction
    {
        public Del(IDictionary<string, object> v, ICentipedeCore core)
            : base("Delete File", v, core)
        { }

        [ActionArgument]
        public String Filename = "";

        [ActionArgument]
        public Boolean SendToRecyclingBin = true;

        protected override void DoAction()
        {
            foreach (FileInfo file in Pwd.EnumerateFiles(ParseStringForVariable(Filename)))
            {
                RecycleOption option = SendToRecyclingBin
                                               ? RecycleOption.SendToRecycleBin
                                               : RecycleOption.DeletePermanently;
                FileSystem.DeleteFile(file.FullName, UIOption.AllDialogs, option);
            }
        }
    }

    [ActionCategory("Shell", DisplayName = "Run Program")]
    public class Start : ShellAction
    {
        [ActionArgument]
        public string Filename;

        [ActionArgument]
        public string Arguments;

        public Start(IDictionary<string, object> v, ICentipedeCore core)
                : base("Run Program", v, core)
        { }

        protected override void DoAction()
        {
            ProcessStartInfo startInfo = new ProcessStartInfo(ParseStringForVariable(Filename), ParseStringForVariable(Arguments))
                                         {
                                                 CreateNoWindow = this.Background,
                                                 RedirectStandardError = true,
                                                 RedirectStandardOutput =   true
                                         };

            Process process = Process.Start(startInfo);

            process.OutputDataReceived += (sender, args) => Message(args.Data, MessageLevel.Action);
            process.ErrorDataReceived += (sender, args) => Message(args.Data, MessageLevel.Error);

            if (!this.DontWait)
            {
                process.WaitForExit();
            }
            
        }

        
        [ActionArgument]
        public bool Background;

        [ActionArgument]
        public bool DontWait;

    }


    [ActionCategory("Shell", DisplayName = "Get Dir from filename")]
    public class GetContainingDirectory : ShellAction
    {
        public GetContainingDirectory(IDictionary<string, object> v, ICentipedeCore core)
                : base("Get Dir from filename", v, core)
        { }

        [ActionArgument]
        public String Filename;

        [ActionArgument(Literal=true)]
        public String DestinationVar = "Path";
        
        protected override void DoAction()
        {
            Variables[DestinationVar] = Path.GetDirectoryName(ParseStringForVariable(Filename));
        }
    }
}
