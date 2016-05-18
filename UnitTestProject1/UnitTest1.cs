using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using entityXML;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Xml.Linq;
using System.IO;
using System.Timers;
namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        string cooperPath = "d:\\cooper";
        string xmlFromPath = @"C:\Users\RDCP01\Desktop";
        string xmlToPath = @"C:\Users\RDCP01\Desktop\move\";

        [TestMethod]
        public void importXML()
        {
            controllerXMLtoFoxpro controlObj = new controllerXMLtoFoxpro();
            controlObj._cooperPath = "d:\\cooper";
            controlObj._xmlFromPath = @"C:\Users\RDCP01\Desktop";
            controlObj._xmlToPath = @"C:\Users\RDCP01\Desktop\move\";
            controlObj._isWriteLog = "Y";
            controlObj.start();
        }

        [TestMethod]
        public void tmmer()
        {
            
            int intInterval = 5;
            Timer MyTimer = new Timer();
            MyTimer.Elapsed += new ElapsedEventHandler(MyTimer_Elapsed);
            MyTimer.Interval = intInterval * (1000 * 60);
            MyTimer.Start();
        }

        private void MyTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            Console.Write(DateTime.Now.ToString("HHmmss"));            
        }

        [TestMethod]
        public void testControlPatient(){

               IList<viewTransformData> patientColumns = new List<viewTransformData>()
                {
                    new viewTransformData(){
                        elementAttribute="P01",
                        foxproField="異動方式14",
                        oledbType = OleDbType.Char
                    },
                    new viewTransformData(){
                        elementAttribute="P02",
                        foxproField="異動日期14",
                        oledbType = OleDbType.Char
                    },
                
                    new viewTransformData(){
                        elementAttribute="P03",
                        foxproField="病歷編號",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="P04",
                        foxproField="病患姓名",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="P05",
                        foxproField="性別",
                        oledbType = OleDbType.Char
                    },
                               
                      new viewTransformData(){
                        elementAttribute="P06",
                        foxproField="出生日期",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="P07",
                        foxproField="身份證號",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="P08",
                        foxproField="住宅電話",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="P09",
                        foxproField="公司電話",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="P10",
                        foxproField="行動電話",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="P11",
                        foxproField="行動電話2",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="P12",
                        foxproField="email",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="P13",
                        foxproField="地址",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="P14",
                        foxproField="初診日期",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="P15",
                        foxproField="部份負擔碼",
                        oledbType = OleDbType.Numeric,
                        value="0",
                    },
                };

            controllerXMLtoFoxpro controlObj = new controllerXMLtoFoxpro();
            controlObj._cooperPath = "d:\\cooper";
            controlObj._xmlFromPath = @"C:\Users\RDCP01\Desktop";
            controlObj._xmlToPath = @"C:\Users\RDCP01\Desktop\move\";
            controlObj.importData("patient.xml", "PATIENT", "PATIENT", "身份證號", patientColumns);
        }

        [TestMethod]
        public void testUpdateXML()
        {
              IList<viewTransformData> viewTransFormDataObjs = new List<viewTransformData>()
            {
                new viewTransformData(){
                    elementAttribute="D01",
                    foxproField="異動方式14",
                    oledbType = OleDbType.Char,                    
                },
                new viewTransformData(){
                    elementAttribute="D02",
                    foxproField="異動日期14",
                    oledbType = OleDbType.Char
                },
                
                new viewTransformData(){
                    elementAttribute="D03",
                    foxproField="醫師代號",
                    oledbType = OleDbType.Char
                },

                  new viewTransformData(){
                    elementAttribute="D04",
                    foxproField="醫師姓名",
                    oledbType = OleDbType.Char
                },

                  new viewTransformData(){
                    elementAttribute="D05",
                    foxproField="身份證號",
                    oledbType = OleDbType.Char
                },
                               
                  new viewTransformData(){
                    elementAttribute="D06",
                    foxproField="合約到期日",
                    oledbType = OleDbType.Char
                },

                  new viewTransformData(){
                    elementAttribute="D07",
                    foxproField="到職日期",
                    oledbType = OleDbType.Char
                },

                  new viewTransformData(){
                    elementAttribute="D08",
                    foxproField="離職日期",
                    oledbType = OleDbType.Char
                },

                  new viewTransformData(){
                    elementAttribute="D09",
                    foxproField="出生日期",
                    oledbType = OleDbType.Char
                },

                  new viewTransformData(){
                    elementAttribute="D10",
                    foxproField="聯絡電話",
                    oledbType = OleDbType.Char
                },

                  new viewTransformData(){
                    elementAttribute="D11",
                    foxproField="行動電話",
                    oledbType = OleDbType.Char
                },

                  new viewTransformData(){
                    elementAttribute="D12",
                    foxproField="員工証號",
                    oledbType = OleDbType.Char
                },
            };


              entityXML.entityXmlToFoxpro entityXMLObj = new entityXML.entityXmlToFoxpro("d:\\cooper");           
              entityXMLObj.ReadFromXML("test.xml", "DOCTOR", viewTransFormDataObjs);              
              //entityXMLObj.EditToFoxpro("doctor",6);
        }

        [TestMethod]
        public void checkDataExist()
        {
            try
            {
                entityXML.entityXmlToFoxpro entityXMLObj = new entityXML.entityXmlToFoxpro("d:\\cooper");                
                int ikey=0;
                bool bolResult=  entityXMLObj.CheckIdentityExist("doctor", "身份證號", "A123456789",out ikey);
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
          


        [TestMethod]
        public void DELETE()
        {
            //entityXmlToFoxpro obj = new entityXmlToFoxpro("d:\\Cooper");
            //obj.DeleteWithZapToFoxpro("rc11");

            int intCurrentCount = 10;
            int intTotalCount = 9;
            if (intCurrentCount >= intTotalCount)
            {
                Assert.AreNotEqual(intCurrentCount, intTotalCount);
            }

        }

        [TestMethod]
        public void importTdata()
        {
            controllerXMLtoFoxpro controlObj = new controllerXMLtoFoxpro();
            controlObj._cooperPath = cooperPath;
            controlObj._xmlFromPath = xmlFromPath;
            string xmlPathFile = Path.Combine(xmlFromPath, "rcxml_10501_送核.xml");
            string orgName = "11111";
            string exspenseYM = "10501";
            string applyType = "11";
            controlObj.importTData(xmlPathFile, out orgName, out exspenseYM, out applyType);
        }

        [TestMethod]
        public void importDdata()
        {
            controllerXMLtoFoxpro controlObj = new controllerXMLtoFoxpro();
            controlObj._cooperPath = cooperPath;
            controlObj._xmlFromPath = xmlFromPath;
            string xmlPathFile = Path.Combine(xmlFromPath, "rcxml_10501_送核.xml");

            string orgName = "11111";
            string exspenseYM = "10501";
            string applyType = "11";
            controlObj.importDdata(xmlPathFile, orgName, exspenseYM, applyType);
        }

        [TestMethod]
        public void customerValue()
        {
            //申報的資料
            string searchPattern = "*.xml";  // This would be for you to construct your prefix
            DirectoryInfo di = new DirectoryInfo(xmlFromPath);
            FileInfo[] files = di.GetFiles(searchPattern);
            foreach (FileInfo fileObj in files)
            {
                if (fileObj.FullName.ToUpper().Contains("DOCTOR"))
                {
                    Console.Write(fileObj.FullName);
                }                
            }

            //int multiplication = 0;
            //int.TryParse("1.0", out multiplication);

            //entityXmlToFoxpro obj = new entityXmlToFoxpro("d:\\Cooper");
            //string value= entityXmlToFoxpro.definitionValue("*:100,ZERO:6", "20.00");
        }
    
    }
}
