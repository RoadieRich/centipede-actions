using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using CentipedeInterfaces;
using Microsoft.VisualBasic.FileIO;
using SearchOption = System.IO.SearchOption;


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
            catch (ArgumentException e)
            {
                throw new ActionException(string.Format("Bad Path {0}", DirectoryPath), e, this);
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
                throw new ActionException(string.Format("Directory not found: {0}", DirectoryName),
                                          e,
                                          this);
            }
            catch (ArgumentException e)
            {
                throw new ActionException(string.Format("Bad Path {0}", DirectoryName), e, this);
            }
            catch (InvalidOperationException e)
            {
                throw new ActionException(string.Format("Bad Path {0}", DirectoryName), e, this);
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

            DirectoryInfo directory = Path.IsPathRooted(@from) ? FileSystem.GetDirectoryInfo(@from) : Pwd;
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

        [ActionArgument]
        public string Verb;

        [ActionArgument]
        public bool Background;

        [ActionArgument]
        public bool DontWait;

        public Start(IDictionary<string, object> v, ICentipedeCore core)
                : base("Run Program", v, core)
        { }

        protected override void DoAction()
        {
            var startInfo =
                new ProcessStartInfo(this.ParseStringForVariable(this.Filename),
                                     this.ParseStringForVariable(this.Arguments))
                {
                    Verb = String.IsNullOrWhiteSpace(this.Verb) ? this.ParseStringForVariable(this.Verb) : null,
                    CreateNoWindow = this.Background,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true
                };

            Process process = Process.Start(startInfo);

            process.OutputDataReceived += (sender, args) => Message(args.Data, MessageLevel.Action);
            process.ErrorDataReceived += (sender, args) => Message(args.Data, MessageLevel.Error);

            if (!this.DontWait)
            {
                process.WaitForExit();
            }
        }
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

    [ActionCategory("Shell", DisplayName = "Get files in directory")]
    public class GetFilesInDirectory : ShellAction
    {
        public GetFilesInDirectory(IDictionary<string, object> v, ICentipedeCore core)
            : base("Get files in directory", v, core)
        { }

        [ActionArgument(Literal = true, DisplayName="Destination var")]
        public string DestinationVariable = "";
        
        [ActionArgument(Literal = false)]
        public string Pattern = "*.*";

        [ActionArgument]
        public bool Recurse = false;

        

        protected override void DoAction()
        {
            var pattern = ParseStringForVariable(Pattern);
            FileInfo[] fileInfos = Pwd.GetFiles(pattern, this.Recurse ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly);
            Variables[DestinationVariable] = fileInfos.Select(fi => fi.FullName).ToList();
        }
    }

    [ActionCategory("Shell", DisplayName = "Print file")]
    public class PrintFile : ShellAction
    {
        public PrintFile(IDictionary<string, object> v, ICentipedeCore core)
            : base("Print file", v, core)
        { }

        [ActionArgument]
        public string Filename = "";
        
        protected override void DoAction()
        {
            Process p = new Process
                        {
                            StartInfo = new ProcessStartInfo
                                        {
                                            CreateNoWindow = true,
                                            Verb = "print",
                                            FileName = ParseStringForVariable(this.Filename)
                                        }
                        };
            p.Start();
        }
    }

    [ActionCategory("Shell", DisplayName = "Open file")]
    public class OpenFile : ShellAction
    {
        public OpenFile(IDictionary<string, object> v, ICentipedeCore core)
            : base("Open file", v, core)
        { }

        [ActionArgument]
        public string Filename = "";

        protected override void DoAction()
        {
            Process p = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    Verb = "open",
                    FileName = ParseStringForVariable(this.Filename)
                }
            };
            p.Start();
        }
    }

}
