using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PythonEngine;
using Rhino.Mocks;
using CentipedeInterfaces;
using Centipede;
using ShellActions;

namespace UnitTestProject1
{
    [TestClass]
    public class SetCwdTests
    {
        [TestMethod]
        public void TestRun()
        {
            var mockCore = MockRepository.GenerateMock<ICentipedeCore>();
            var mockVariables = MockRepository.GenerateMock<IDictionary<string, object>>();


            var oldPwd = ShellActionPwdAccessor.AccessedPwd;

            var action = new SetCwd(mockVariables, mockCore) { DirectoryPath = ".." };

            action.Run();

            Assert.AreNotEqual(oldPwd, ShellActionPwdAccessor.AccessedPwd);
            ShellActionPwdAccessor.AccessedPwd = oldPwd;
        }

        [TestMethod]
        public void TestInterpolation()
        {

            var mockCore = MockRepository.GenerateMock<ICentipedeCore>();
            var mockVariables = MockRepository.GenerateMock<IDictionary<string, object>>();
            var mockPythonEngine = MockRepository.GenerateStub<IPythonEngine>();


            var oldPwd = ShellActionPwdAccessor.AccessedPwd;

            mockCore.Expect(c => c.PythonEngine).Return(mockPythonEngine);

            var bc = MockRepository.GenerateMock<IPythonByteCode>();

            mockPythonEngine.Expect(
                                    e => e.Compile(Arg<String>.Is.Equal("'..'"),
                                              Arg<SourceCodeType>.Is.Anything))
                            .Return(bc);
            mockPythonEngine.Expect(e => e.Evaluate(bc)).Return("..");


            var action = new SetCwd(mockVariables, mockCore) { DirectoryPath = @"{'..'}" };

            action.Run();

            mockPythonEngine.VerifyAllExpectations();

            ShellActionPwdAccessor.AccessedPwd = oldPwd;
        }



        [TestMethod]
        public void TestInvalidPath()
        {
            var mockCore = MockRepository.GenerateMock<ICentipedeCore>();
            var mockVariables = MockRepository.GenerateMock<IDictionary<string, object>>();


            var oldPwd = ShellActionPwdAccessor.AccessedPwd;

            var action = new SetCwd(mockVariables, mockCore)
                            {
                                DirectoryPath = "!£$%^*&()&" //no swearing in the source code!
                            };

            try
            {
                action.Run();
                Assert.Fail("Did not throw!");
            }
            catch (Exception e)
            {
                Assert.IsInstanceOfType(e, typeof(ActionException));
            }

            ShellActionPwdAccessor.AccessedPwd = oldPwd;


        }
    }

    internal abstract class ShellActionPwdAccessor : ShellAction
    {
        private ShellActionPwdAccessor()
            : base(null, null, null)
        { }

        public static DirectoryInfo AccessedPwd
        {
            get { return Pwd; }
            set { Pwd = value; }
        }
    }

    [TestClass]
    public class CdTests
    {
        [TestMethod]
        public void TestRun()
        {
            const string dirName = "adir";



            var mockCore = MockRepository.GenerateMock<ICentipedeCore>();
            var mockVariables = MockRepository.GenerateMock<IDictionary<string, object>>();

            using (new DirForTesting(dirName))
            {
                var oldPwd = ShellActionPwdAccessor.AccessedPwd;

                var action = new Cd(mockVariables, mockCore) { DirectoryName = dirName };

                action.Run();

                Assert.AreNotEqual(oldPwd, ShellActionPwdAccessor.AccessedPwd);

                ShellActionPwdAccessor.AccessedPwd = oldPwd;
            }
        }

        [TestMethod]
        public void TestInterpolation()
        {

            var oldPwd = ShellActionPwdAccessor.AccessedPwd;
            const string dirname = "anotherDir";
            const string dirtext = "<placeholder>";

            var mockCore = MockRepository.GenerateMock<ICentipedeCore>();
            var mockVariables = MockRepository.GenerateMock<IDictionary<string, object>>();
            var mockPythonEngine = MockRepository.GenerateStub<IPythonEngine>();
            
            mockCore.Expect(c => c.PythonEngine).Return(mockPythonEngine);

            var bc = MockRepository.GenerateMock<IPythonByteCode>();

            mockPythonEngine.Expect(
                                    e => e.Compile(Arg<String>.Is.Equal(dirtext),
                                              Arg<SourceCodeType>.Is.Anything))
                            .Return(bc);
            mockPythonEngine.Expect(e => e.Evaluate(bc)).Return(dirname);


            var action = new Cd(mockVariables, mockCore)
                         {
                             DirectoryName = "{" + dirtext + "}"
                         };
            using (new DirForTesting(dirname))
            {
                action.Run();
            }


            mockPythonEngine.VerifyAllExpectations();

            ShellActionPwdAccessor.AccessedPwd = oldPwd;
        }

        [TestMethod]
        public void TestInvalidPath()
        {
            var oldPwd = ShellActionPwdAccessor.AccessedPwd;
            var mockCore = MockRepository.GenerateMock<ICentipedeCore>();
            var mockVariables = MockRepository.GenerateMock<IDictionary<string, object>>();

            var action = new Cd(mockVariables, mockCore)
            {
                DirectoryName = "!£$%^*&()&" //no swearing in the source code!
            };

            try
            {
                action.Run();
                Assert.Fail("Did not throw!");
            }
            catch (Exception e)
            {
                Assert.IsInstanceOfType(e, typeof(ActionException));
            }

            ShellActionPwdAccessor.AccessedPwd = oldPwd;
        }

        [TestMethod]
        public void TestMissingDirectory()
        {
            var oldPwd = ShellActionPwdAccessor.AccessedPwd;
            var mockCore = MockRepository.GenerateMock<ICentipedeCore>();
            var mockVariables = MockRepository.GenerateMock<IDictionary<string, object>>();

            var action = new Cd(mockVariables, mockCore)
            {
                DirectoryName = "a\\directory\\that\\doesn.t\\exist"
            };

            try
            {
                action.Run();
                Assert.Fail("Did not throw!");
            }
            catch (Exception e)
            {
                Assert.IsInstanceOfType(e, typeof(ActionException));
            }

            ShellActionPwdAccessor.AccessedPwd = oldPwd;
        }
    }

    [TestClass]
    public class TestCp
    {
        [TestMethod]
        public void TestRun()
        {

            var mockCore = MockRepository.GenerateMock<ICentipedeCore>();
            var mockVariables = MockRepository.GenerateMock<IDictionary<string, object>>();

            const string fromName = "from";
            const string toName = "to";

            Cp action = new Cp(mockVariables, mockCore)
                        {
                            From = fromName,
                            To = toName
                        };
            const int length = 1024;
            using (new FileForTesting(fromName, length))
            using (new FileForTesting(toName, create: false))
            {
                action.Run();

                Assert.IsTrue(File.Exists(toName));
                Assert.AreEqual(File.ReadAllText(toName).Length, length);
            }
        }

        [TestMethod]
        public void TestInterpolation()
        {

            const string fromName = "from1";
            const string fromText = "<from>";
            const string toName = "to2";
            const string toText = "<to>";

            var mockCore = MockRepository.GenerateMock<ICentipedeCore>();
            var mockVariables = MockRepository.GenerateMock<IDictionary<string, object>>();
            var mockPythonEngine = MockRepository.GenerateStub<IPythonEngine>();

            mockCore.Expect(c => c.PythonEngine).Return(mockPythonEngine);

            var bcFrom = MockRepository.GenerateMock<IPythonByteCode>();
            var bcTo = MockRepository.GenerateMock<IPythonByteCode>();


            mockPythonEngine.Expect(e => e.Compile(Arg<String>.Is.Equal(fromText),
                                                   Arg<SourceCodeType>.Is.Anything))
                            .Return(bcFrom);

            mockPythonEngine.Expect(e => e.Evaluate(bcFrom))
                            .Return(fromName);

            mockPythonEngine.Expect(e => e.Compile(Arg<String>.Is.Equal(toText),
                                                   Arg<SourceCodeType>.Is.Anything))
                            .Return(bcTo);

            mockPythonEngine.Expect(e => e.Evaluate(bcTo)).Return(toName);

            var action = new Cp(mockVariables, mockCore)
                         {
                             From = string.Format("{{{0}}}", fromText),
                             To = string.Format("{{{0}}}", toText)
                         };

            using (new FileForTesting(fromName))
            using (new FileForTesting(toName, create: false))
            {
                action.Run();
            }

            mockPythonEngine.VerifyAllExpectations();
        }

        [TestMethod]
        public void TestOverwrite()
        {

            var mockCore = MockRepository.GenerateMock<ICentipedeCore>();
            var mockVariables = MockRepository.GenerateMock<IDictionary<string, object>>();

            const string fromName = "fromow";
            const string toName = "toow";

            Cp action = new Cp(mockVariables, mockCore)
                        {
                            From = fromName,
                            To = toName,
                            AllowOverWrite = false
                        };

            using (new FileForTesting(fromName))
            using (new FileForTesting(toName))
            {
                try
                {
                    action.Run();
                    Assert.Fail("Did not fail when not overwriting");
                }
                catch
                { }
                action.AllowOverWrite = true;
                try
                {
                    action.Run();
                }
                catch
                {
                    Assert.Fail("Failed to overwrite");
                }
            }
        }

        [TestMethod]
        public void TestMissingFile()
        {
            var mockCore = MockRepository.GenerateMock<ICentipedeCore>();
            var mockVariables = MockRepository.GenerateMock<IDictionary<string, object>>();

            const string fromName = "fromow";
            const string toName = "toow";

            Cp action = new Cp(mockVariables, mockCore)
                        {
                            From = fromName,
                            To = toName,
                            AllowOverWrite = false
                        };
            try
            {
                action.Run();
                Assert.Fail("Did not fail with missing file");
            }
            catch
            { }
        }
    }


    [TestClass]
    public class MkDirTests
    {
        [TestMethod]
        public void TestRun()
        {
            
        }
    }

    
    class DirForTesting : IDisposable
    {
        public DirForTesting(string name, bool create = true)
        {
            this._dir = new DirectoryInfo(Path.Combine(Environment.CurrentDirectory, name));
            if (create)
            {
                this._dir.Create();
            }
        }

        private readonly DirectoryInfo _dir;

        public void Dispose()
        {
            this._dir.Delete(true);
        }
    }

    class FileForTesting : IDisposable
    {

        public FileForTesting(string name, int length = 0, bool create = true)
        {
            this._file = new FileInfo(name);
            if (create)
            {
                using (var stream = _file.Open(FileMode.OpenOrCreate))
                {
                    for (int i = 0; i < length; i++)
                    {
                        stream.WriteByte(0x61);
                    }
                }
            }
        }

        private readonly FileInfo _file;
        

        public void Dispose()
        {
            try
            {
                this._file.Delete();
            }
            catch (IOException e)
            {
                
            }
        }
    }
}