using ClassLibraryFoxDB;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using WriteEvent;

namespace entityXML
{
    public class controllerXMLtoFoxpro
    {
        /// <summary>
        /// 傳遞XML和FoxproDB的對照檔
        /// </summary>
        public IList<viewTransformData> _listViewTransformDataObjs
        {
            get;
            set;
        }

        /// <summary>
        /// XML的路徑
        /// </summary>
        public string _xmlFromPath
        {
            get;
            set;
        }

        /// <summary>
        /// 搬移指定的目錄
        /// </summary>
        public string _xmlToPath
        {
            get;
            set;
        }

        /// <summary>
        /// cooper的路徑
        /// </summary>
        public string _cooperPath
        {
            get;
            set;
        }

        /// <summary>
        /// 是否要寫日誌檔
        /// </summary>
        public string _isWriteLog
        {
            get;
            set;
        }

        WrittingEventLog _writeObj = new WrittingEventLog();

        private string _currentPath = Directory.GetCurrentDirectory();

        public void start()
        {
            try
            {
                if (_isWriteLog.ToUpper() == "Y")
                {
                    _writeObj.writeToFile("開始start");
                }

                //申報的資料
                string searchPattern = "*.xml";  // This would be for you to construct your prefix
                DirectoryInfo di = new DirectoryInfo(_xmlFromPath);
                FileInfo[] files = di.GetFiles(searchPattern);

                if (_isWriteLog.ToUpper() == "Y")
                {
                    string json = JsonConvert.SerializeObject(files);
                    _writeObj.writeToFile("偵測檔案:"+json);
                }

                string orgName = "";
                string ymDate="";
                string applyType="";
                foreach (FileInfo fileObj in files)
                {
                    if (fileObj.FullName.ToUpper().Contains("DOCTOR"))
                    {
                        if (_isWriteLog.ToUpper() == "Y")
                            _writeObj.writeToFile("匯入Doctor資料表");

                        #region doctorColumns
                        IList<viewTransformData> doctorColumns = new List<viewTransformData>()
                        {
                            new viewTransformData(){
                                elementAttribute="D01",
                                foxproField="異動方式14",
                                oledbType = OleDbType.Char
                            },
                            new viewTransformData(){
                                elementAttribute="D02",
                                foxproField="異動日期14",
                                oledbType = OleDbType.Char
                            },
                
                            new viewTransformData(){
                                elementAttribute="D03",
                                foxproField="員工証號",
                                oledbType = OleDbType.Char,
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
                                elementAttribute="D03",
                                foxproField="醫師代號",
                                oledbType = OleDbType.Char
                            },
                        };

                        #endregion

                        importData(fileObj.FullName, "DOCTOR", "doctor", "身份證號", doctorColumns);
                    }
                    else if (fileObj.FullName.ToUpper().Contains("PATIENT"))
                    {
                        #region patientColumns

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
                                oledbType = OleDbType.Numeric
                            },
                        };

                        #endregion

                        importData(fileObj.FullName, "PATIENT", "patient", "身份證號", patientColumns);                      
                    }
                    else if (fileObj.FullName.ToUpper().Contains("RCXML"))
                    {
                        entityXmlToFoxpro entityXMLObj = new entityXmlToFoxpro(_cooperPath);

                        entityXMLObj.DeleteWithZapToFoxpro("RC10"); //zap清空資料
                        entityXMLObj.DeleteWithZapToFoxpro("RC11"); //zap清空資料
                        entityXMLObj.DeleteWithZapToFoxpro("RC12"); //zap清空資料
                        //匯入RC10

                        importTData(fileObj.FullName, out orgName, out ymDate, out applyType);
                        
                        //匯入RC11、RC12
                        importDdata(fileObj.FullName, orgName, ymDate, applyType);
                    }

                    MoveFile(fileObj.Name, _xmlFromPath, _xmlToPath);
                }
            }
            catch (Exception ex)
            {
                //throw new Exception(ex.Message);
                _writeObj.writeToFile(DateTime.Now.ToString("yyyyMMdd") + "_errorLog", _currentPath, "start: " + ex.Message.ToString());
            }
        }

        public void importData(string xmlFilePathName,string elementName,string tableName, string checkRepeatColumn, IList<viewTransformData> listViewTransformDataObjs)
        {
            try 
	        {
                /*
                * 1. 檢查目錄中是否有已有xml 檔              
                * 2. 如果有檔案逐一匯入資料表中 
                * 3。匯入完畢後，儲存至另一個目錄
                */
                if (File.Exists(xmlFilePathName))
                {
                    entityXmlToFoxpro entityXMLObj = new entityXmlToFoxpro(_cooperPath);
                    Dictionary<int, IList<viewTransformData>> dicDataXMLRows = entityXMLObj.ReadFromXML(xmlFilePathName, elementName, listViewTransformDataObjs);                    
                    string strIdentity = "";                    
                    foreach (KeyValuePair<int, IList<viewTransformData>> entry in dicDataXMLRows)
                    {
                        // do something with entry.Value or entry.Key
                        strIdentity = (from q in entry.Value
                                       where q.foxproField == checkRepeatColumn
                                       select q.value).FirstOrDefault();
                        
                        if(tableName.ToUpper()=="DOCTOR")
                            importDoctor(tableName, checkRepeatColumn, entityXMLObj, strIdentity, entry);
                         else
                            importPatient(tableName, checkRepeatColumn, entityXMLObj, strIdentity, entry);
                    }
                }
	        }
	        catch (Exception ex)
	        {
                _writeObj.writeToFile(DateTime.Now.ToString("yyyyMMdd") + "_errorLog", _currentPath, "checkExistDoctor: 資料表:" + elementName + " 訊息: " +  ex.Message.ToString());          
	        }
        }

        private void importPatient(string tableName, string checkRepeatColumn, entityXmlToFoxpro entityXMLObj, string strIdentity, KeyValuePair<int, IList<viewTransformData>> entry)
        {

            int ikey = 0;
           bool bolExistIdentity = entityXMLObj.CheckIdentityExist(tableName, checkRepeatColumn, strIdentity, out ikey);

            //資料已有存在此身份證
            if (bolExistIdentity == true)
            {
                int intIdentity = 0;
                int.TryParse(strIdentity, out intIdentity);

                if (_isWriteLog.ToUpper() == "Y")
                {
                    string json = JsonConvert.SerializeObject(entry.Value);
                    _writeObj.writeToFile("修改資料:" + checkRepeatColumn + ":" + strIdentity + "<<<" + json + ">>>");
                }

                entityXMLObj.EditToFoxpro(tableName, entry.Value, ikey);
            }
            else
            {
                if (_isWriteLog.ToUpper() == "Y")
                {
                    string json = JsonConvert.SerializeObject(entry.Value);
                    _writeObj.writeToFile("新增資料:" + checkRepeatColumn + ":" + strIdentity + "<<<" + json + ">>>");
                }
                entityXMLObj.AddToFoxpro(tableName, entry.Value);
            }
        }

        private void importDoctor(string tableName, string checkRepeatColumn, entityXmlToFoxpro entityXMLObj, string strIdentity, KeyValuePair<int, IList<viewTransformData>> entry)
        {
            int ikey = 0;
            //醫師的身份證有可能重複，所以另外判斷離職日是否空值
            bool bolExistIdentity = entityXMLObj.CheckDoctorIdentityExist(tableName, checkRepeatColumn, strIdentity, out ikey);

            //資料已有存在此身份證
            if (bolExistIdentity == true)
            {
                if (_isWriteLog.ToUpper() == "Y")
                {
                    string json = JsonConvert.SerializeObject(entry.Value);
                    _writeObj.writeToFile("修改資料:" + checkRepeatColumn + ":" + ikey + "<<<" + json + ">>>");
                }
                entityXMLObj.EditToFoxpro(tableName, entry.Value, ikey);
            }
            else
            {
                if (_isWriteLog.ToUpper() == "Y")
                {
                    string json = JsonConvert.SerializeObject(entry.Value);
                    _writeObj.writeToFile("新增資料:" + checkRepeatColumn + ":" + ikey + "<<<" + json + ">>>");
                }

                //特別要將員代號的值改為流程號編輯
                string strNewDoctorID = entityXMLObj.getNewDoctorID("DOCTOR", "醫師代號");

                var doctorID = entry.Value.Where(q => q.foxproField == "醫師代號").FirstOrDefault();
                doctorID.value = strNewDoctorID;
                entityXMLObj.AddToFoxpro(tableName, entry.Value);
            }
        }

        public void importTData(string xmlFilePathName, out string strOrganizationName, out string strExpenseYearMonth, out string strApplyType)
        {
            try
            {
                entityXmlToFoxpro entityXMLObj = new entityXmlToFoxpro(_cooperPath);
                XElement root = XElement.Load(xmlFilePathName);
                IEnumerable<XElement> xmlTdata = from xmldata in root.Elements("tdata") select xmldata;

                var orgNameObj = xmlTdata.Elements("t2");
                var applyYMObj = xmlTdata.Elements("t3");
                var applyTypeObj = xmlTdata.Elements("t4");

                strOrganizationName = (orgNameObj == null) ? "" : orgNameObj.FirstOrDefault().Value;

                strExpenseYearMonth = (applyYMObj == null) ? "" : applyYMObj.FirstOrDefault().Value;

                strApplyType = (applyTypeObj == null) ? "" : applyTypeObj.FirstOrDefault().Value;

                #region 定義欄位定義 TData
                IList<viewTransformData> listTdata = new List<viewTransformData>(){                    
                        new viewTransformData() {
                        elementAttribute = "t1",
                        foxproField = "資料格式",
                        oledbType = OleDbType.Char
                        },
                        new viewTransformData() {
                            elementAttribute = "t2",
                            foxproField = "機構代號",
                            oledbType = OleDbType.Char
                        },
                         new viewTransformData()
                         {
                             elementAttribute = "t3",
                             foxproField = "費用年月",
                             oledbType = OleDbType.Char
                         },
                          new viewTransformData()
                          {
                              elementAttribute = "t4",
                              foxproField = "申報方式",
                              oledbType = OleDbType.Char
                          },
                          new viewTransformData()
                          {
                              elementAttribute = "t5",
                              foxproField = "申報類別",
                              oledbType = OleDbType.Char
                          },
                      new viewTransformData()
                      {
                          elementAttribute = "t6",
                          foxproField = "申報日期",
                          oledbType = OleDbType.Char
                      },
                      new viewTransformData()
                      {
                          elementAttribute = "t19",
                          foxproField = "牙醫一般件",
                          oledbType = OleDbType.Char
                      },
                       new viewTransformData()
                       {
                           elementAttribute = "t20",
                           foxproField = "牙醫一般額",
                           oledbType = OleDbType.Char
                       },
                       new viewTransformData()
                       {
                           elementAttribute = "t21",
                           foxproField = "牙醫專案件",
                           oledbType = OleDbType.Char
                       },
                        new viewTransformData()
                        {
                            elementAttribute = "t22",
                            foxproField = "牙醫專案額",
                            oledbType = OleDbType.Char
                        },
                        new viewTransformData()
                        {
                            elementAttribute = "t23",
                            foxproField = "牙醫總件數",
                            oledbType = OleDbType.Char
                        },
                        new viewTransformData()
                        {
                            elementAttribute = "t24",
                            foxproField = "牙醫總金額",
                            oledbType = OleDbType.Char
                        },
                         new viewTransformData()
                         {
                             elementAttribute = "t31",
                             foxproField = "預防保健件",
                             oledbType = OleDbType.Char
                         },
                          new viewTransformData()
                          {
                              elementAttribute = "t32",
                              foxproField = "預防保健額",
                              oledbType = OleDbType.Char
                          },
                           new viewTransformData()
                           {
                               elementAttribute = "t37",
                               foxproField = "件數總計",
                               oledbType = OleDbType.Char
                           },
                            new viewTransformData()
                           {
                               elementAttribute = "t38",
                               foxproField = "金額總計",
                               oledbType = OleDbType.Char
                           },
                             new viewTransformData()
                           {
                               elementAttribute = "t41",
                               foxproField = "申報起日期",
                               oledbType = OleDbType.Char
                           },
                              new viewTransformData()
                           {
                               elementAttribute = "t42",
                               foxproField = "申報迄日期",
                               oledbType = OleDbType.Char
                           }};
                #endregion

                IList<viewTransformData> liViewRc10Data = entityXMLObj.readFromXML(xmlTdata, listTdata);

                if (_isWriteLog.ToUpper() == "Y")
                {
                    string json = JsonConvert.SerializeObject(liViewRc10Data);
                    _writeObj.writeToFile("新增TData資料: <<<" + json + ">>>");
                }
                entityXMLObj.AddToFoxpro("RC10", liViewRc10Data,null, false); 
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        /// <summary>
        /// 包含RC11,RC12的資料
        /// </summary>
        /// <param name="xmlFilePathName"></param>
        /// <param name="organizationName"></param>
        /// <param name="expenseYearMonth"></param>
        /// <param name="applyType"></param>
        public void importDdata(string xmlFilePathName, string organizationName, string expenseYearMonth, string applyType)
        {

            entityXmlToFoxpro entityXMLObj = new entityXmlToFoxpro(_cooperPath);
            XElement root = XElement.Load(xmlFilePathName);
            IEnumerable<XElement> xmlAttibutes = from xmldata in root.Elements("ddata") select xmldata;

            IList<viewTransformData> liViewData = new List<viewTransformData>();
            
            //第一層data
            foreach (var ddata in xmlAttibutes)
            {
                #region  第二層 thead  
              
                var headElement = from headObj in ddata.Elements("dhead")
                               select headObj;
                var objCaseType = headElement.Elements("d1");//案件分類
                var objFlowNumber = headElement.Elements("d2"); //流水號
                string strCaseType = (objCaseType == null) ? "" : objCaseType.FirstOrDefault().Value.ToString();
                string strFlowNumber = (objFlowNumber == null) ? "" : objFlowNumber.FirstOrDefault().Value.ToString(); 

                #endregion

                #region 第二層 tbody

                #region 欄位定義
                IList<viewTransformData> listDdata = new List<viewTransformData>(){
                    new viewTransformData(){
                        elementAttribute="d3",
                        foxproField="身份證號",
                        oledbType = OleDbType.Char
                    },
                    new viewTransformData(){
                        elementAttribute="d4",
                        foxproField="治療代號一",
                        oledbType = OleDbType.Char
                    },
                
                    new viewTransformData(){
                        elementAttribute="d5",
                        foxproField="治療代號二",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="d6",
                        foxproField="治療代號三",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="d7",
                        foxproField="治療代號四",
                        oledbType = OleDbType.Char
                    },
                               
                      new viewTransformData(){
                        elementAttribute="d8",
                        foxproField="就醫科別",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="d9",
                        foxproField="就醫日期",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="d9",
                        foxproField="結束日期",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="d10",
                        foxproField="結束日期",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="d11",
                        foxproField="出生日期",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="d12",
                        foxproField="補報註記",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="d14",
                        foxproField="給付類別",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="d15",
                        foxproField="部份負擔號",
                        oledbType = OleDbType.Char
                    },

                       new viewTransformData(){
                        elementAttribute="d17",
                        foxproField="轉入醫院號",
                        oledbType = OleDbType.Char
                    },

                       new viewTransformData(){
                        elementAttribute="d18",
                        foxproField="是否轉出",
                        oledbType = OleDbType.Char
                    },

                       new viewTransformData(){
                        elementAttribute="d19",
                        foxproField="國際病碼一",
                        oledbType = OleDbType.Char
                    },

                        new viewTransformData(){
                        elementAttribute="d20",
                        foxproField="國際病碼二",
                        oledbType = OleDbType.Char
                    },

                       new viewTransformData(){
                        elementAttribute="d21",
                        foxproField="國際病碼三",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="d22",
                        foxproField="國際病碼四",
                        oledbType = OleDbType.Char
                    },                
                      new viewTransformData(){
                        elementAttribute="d23",
                        foxproField="國際病碼五",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="d24",
                        foxproField="主手術代碼",
                        oledbType = OleDbType.Char
                    },

                       new viewTransformData(){
                        elementAttribute="d25",
                        foxproField="次手術碼一",
                        oledbType = OleDbType.Char
                    },

                     new viewTransformData(){
                        elementAttribute="d26",
                        foxproField="次手術碼二",
                        oledbType = OleDbType.Char
                    },

                     new viewTransformData(){
                        elementAttribute="d27",
                        foxproField="給藥日份",
                        oledbType = OleDbType.Char,
                        value="ZERO:2"
                    },

                    new viewTransformData(){
                        elementAttribute="d28",
                        foxproField="調劑方式",
                        oledbType = OleDbType.Char
                    },

                    new viewTransformData(){
                        elementAttribute="d29",
                        foxproField="健保卡序號",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="d30",
                        foxproField="醫師代號",
                        oledbType = OleDbType.Char
                    },

                      new viewTransformData(){
                        elementAttribute="d31",
                        foxproField="藥師代號",
                        oledbType = OleDbType.Char
                    },

                     new viewTransformData(){
                        elementAttribute="d32",
                        foxproField="用藥金額",
                        oledbType = OleDbType.Char,
                        value="ZERO:8"
                    },

                     new viewTransformData(){
                        elementAttribute="d33",
                        foxproField="診療金額",
                        oledbType = OleDbType.Char,
                        value="ZERO:8"
                    },

                     new viewTransformData(){
                        elementAttribute="d34",
                        foxproField="特材金額",
                        oledbType = OleDbType.Char,
                        value="ZERO:8"
                    },

                       new viewTransformData(){
                        elementAttribute="d35",
                        foxproField="診察費代號",
                        oledbType = OleDbType.Char
                    },
                       new viewTransformData(){
                        elementAttribute="d36",
                        foxproField="診察費",
                        oledbType = OleDbType.Char,
                        value="ZERO:8"
                    },

                       new viewTransformData(){
                        elementAttribute="d37",
                        foxproField="藥事代號",
                        oledbType = OleDbType.Char
                    },

                       new viewTransformData(){
                        elementAttribute="d38",
                        foxproField="藥事服務費",
                        oledbType = OleDbType.Char,
                        value="ZERO:8"
                    },
                        new viewTransformData(){
                        elementAttribute="d39",
                        foxproField="合計金額",
                        oledbType = OleDbType.Char,
                        value="ZERO:8"
                    },

                     new viewTransformData(){
                        elementAttribute="d40",
                        foxproField="部份負擔額",
                        oledbType = OleDbType.Char,
                        value="ZERO:8"
                    },

                      new viewTransformData(){
                        elementAttribute="d41",
                        foxproField="申請金額",
                        oledbType = OleDbType.Char,
                        value="ZERO:8"
                    },

                     new viewTransformData(){
                        elementAttribute="d42",
                        foxproField="論病例計酬",
                        oledbType = OleDbType.Char
                    },

                     new viewTransformData(){
                        elementAttribute="d43",
                        foxproField="代辦費金額",
                        oledbType = OleDbType.Char,
                        value="ZERO:6"
                    },

                     new viewTransformData(){
                        elementAttribute="d45",
                        foxproField="新生兒生日",
                        oledbType = OleDbType.Char
                    },

                     new viewTransformData(){
                        elementAttribute="d46",
                        foxproField="急診開始時",
                        oledbType = OleDbType.Char
                    },

                     new viewTransformData(){
                        elementAttribute="d47",
                        foxproField="急診結束時",
                        oledbType = OleDbType.Char
                    },

                     new viewTransformData(){
                        elementAttribute="d48",
                        foxproField="給付提升碼",
                        oledbType = OleDbType.Char
                    },

                     new viewTransformData(){
                        elementAttribute="d49",
                        foxproField="姓名",
                        oledbType = OleDbType.Char
                    },

                    new viewTransformData(){
                        elementAttribute="d50",
                        foxproField="矯正機關號",
                        oledbType = OleDbType.Char
                    },
                     new viewTransformData(){
                        elementAttribute="d52",
                        foxproField="",
                        oledbType = OleDbType.Char
                    },
                     new viewTransformData(){
                        elementAttribute="d53",
                        foxproField="",
                        oledbType = OleDbType.Char
                    }
                }; 
                #endregion

                var bodyElement = from bodyObj in ddata.Elements("dbody")
                                  select bodyObj;

                var objd52 = bodyElement.Elements("d52").FirstOrDefault();//案件分類
                var objd53 = bodyElement.Elements("d53").FirstOrDefault(); //流水號

                var objIdentity = bodyElement.Elements("d3").FirstOrDefault();//身份證

                string strd52 = (objd52 == null) ? "  " :objd52.Value.ToString().PadLeft(2,' ');
                string strd53 = (objd53 == null) ? "    " :objd53.Value.ToString().PadLeft(4, ' ');
                string strIdentity = (objIdentity == null) ? "" : objIdentity.Value.ToString();

                IList<viewTransformData> liViewRc11Data = entityXMLObj.readFromXML(bodyElement, listDdata);

                #region 額外新增RC11的欄位
                List<columnsData> liAdditionColumns = new List<columnsData>(){
                    new columnsData()
                    {
                        strFileName = "醫事機構號",
                        oledbTypeValue = OleDbType.Char,
                        strValue = organizationName
                    },
                    new columnsData()
                    {
                        strFileName = "費用年月",
                        oledbTypeValue = OleDbType.Char,
                        strValue = expenseYearMonth
                    },
                     new columnsData()
                    {
                        strFileName = "申報類別",
                        oledbTypeValue = OleDbType.Char,
                        strValue = applyType
                    },
                      new columnsData()
                    {
                        strFileName = "資料格式",
                        oledbTypeValue = OleDbType.Char,
                        strValue = "11"
                    },
                     new columnsData()
                    {
                        strFileName = "連續總日份",
                        oledbTypeValue = OleDbType.Char,
                        strValue = "00"
                    },
                      new columnsData()
                    {
                        strFileName = "R11MARK",
                        oledbTypeValue = OleDbType.Char,
                        strValue = strd52+strd53
                    },
                
                    new columnsData()
                    {
                        strFileName = "案件分類",
                        oledbTypeValue = OleDbType.Char,
                        strValue = strCaseType
                    },
                    new columnsData()
                    {
                        strFileName = "流水號",
                        oledbTypeValue = OleDbType.Char,
                        strValue = strFlowNumber
                    },                    
                }; 
                #endregion

                #region 儲存RC11

                if (_isWriteLog.ToUpper() == "Y")
                {
                    string json = JsonConvert.SerializeObject(liViewRc11Data);
                    _writeObj.writeToFile("新增DTata.Rc11資料: <<<" + json + ">>>");
                }

                entityXMLObj.AddToFoxpro("RC11", liViewRc11Data, liAdditionColumns, false); 
                #endregion

                #endregion

                #region 第三層pdata

                #region listPDatas 定義pData的欄位
                IList<viewTransformData> listPDatas = new List<viewTransformData>(){
                           new viewTransformData(){
                                elementAttribute="p2",
                                foxproField="調劑方式",
                                oledbType = OleDbType.Char
                            },
                              new viewTransformData(){
                                elementAttribute="p4",
                                foxproField="藥品代號",
                                oledbType = OleDbType.Char,
                            },
                              new viewTransformData(){
                                elementAttribute="p5",
                                foxproField="藥品用量",
                                oledbType = OleDbType.Char,
                                value="*:100,ZERO:6"
                            },
                             new viewTransformData(){
                                elementAttribute="p6",
                                foxproField="藥品用量",
                                oledbType = OleDbType.Char,
                            },
                              new viewTransformData(){
                                elementAttribute="p7",
                                foxproField="使用頻率",
                                oledbType = OleDbType.Char,
                            },
                               new viewTransformData(){
                                elementAttribute="p8",
                                foxproField="使用頻率",
                                oledbType = OleDbType.Char,
                            },
                               new viewTransformData(){
                                elementAttribute="p9",
                                foxproField="給藥途徑",
                                oledbType = OleDbType.Char,
                            },
                              new viewTransformData(){
                                elementAttribute="p10",
                                foxproField="總量",
                                oledbType = OleDbType.Char,
                                value="*:10,ZERO:6"
                            },
                             new viewTransformData(){
                                elementAttribute="p11",
                                foxproField="單價",
                                oledbType = OleDbType.Char,
                                value="*:100,ZERO:9"
                            },
                             new viewTransformData(){
                                elementAttribute="p12",
                                foxproField="金額",
                                oledbType = OleDbType.Char,
                                value="ZERO:8"
                            },
                              new viewTransformData(){
                                elementAttribute="p13",
                                foxproField="醫令序",
                                oledbType = OleDbType.Char
                            },
                             new viewTransformData(){
                                  elementAttribute="p14",
                                  foxproField="執行起時",
                                  oledbType = OleDbType.Char
                              },
                               new viewTransformData(){
                                  elementAttribute="p15",
                                  foxproField="執行迄時",
                                  oledbType = OleDbType.Char
                              },
                                new viewTransformData(){
                                  elementAttribute="p16",
                                  foxproField="醫事人號",
                                  oledbType = OleDbType.Char
                              },

                                new viewTransformData(){
                                  elementAttribute="p17",
                                  foxproField="案件註記",
                                  oledbType = OleDbType.Char
                              },

                                new viewTransformData(){
                                  elementAttribute="p18",
                                  foxproField="影像來源",
                                  oledbType = OleDbType.Char
                              },

                                new viewTransformData(){
                                  elementAttribute="p19",
                                  foxproField="事前審號",
                                  oledbType = OleDbType.Char
                              },

                                new viewTransformData(){
                                  elementAttribute="p20",
                                  foxproField="就醫科別",
                                  oledbType = OleDbType.Char
                                },
                };
                #endregion

                #region 另自訂增加RC12欄位
                List<columnsData> liAddRC12Columns = new List<columnsData>(){
                    new columnsData()
                    {
                        strFileName = "醫事機構號",
                        oledbTypeValue = OleDbType.Char,
                        strValue = organizationName
                    },
                    new columnsData()
                    {
                        strFileName = "費用年月",
                        oledbTypeValue = OleDbType.Char,
                        strValue = expenseYearMonth
                    },
                     new columnsData()
                    {
                        strFileName = "申報類別",
                        oledbTypeValue = OleDbType.Char,
                        strValue = applyType
                    },
                      new columnsData()
                    {
                        strFileName = "資料格式",
                        oledbTypeValue = OleDbType.Char,
                        strValue = "12"
                    },
                    new columnsData()
                    {
                        strFileName = "案件分類",
                        oledbTypeValue = OleDbType.Char,
                        strValue = strCaseType
                    },
                    new columnsData()
                    {
                        strFileName = "流水號",
                        oledbTypeValue = OleDbType.Char,
                        strValue = strFlowNumber
                    },    
                     new columnsData()
                    {
                        strFileName = "身份證號",
                        oledbTypeValue = OleDbType.Char,
                        strValue = strIdentity
                    }
                };  
                #endregion

                #region 取得rc12 的xml

                //宣告變數 rc12最多只有五筆處置，所以超過需要另新增一筆
                IList<viewTransformData> foxproRC12Columns = new List<viewTransformData>();
                IList<viewTransformData> foxproRC12_SecondColumns = new List<viewTransformData>();
                IList<viewTransformData> foxproRC12_ThirdColumns = new List<viewTransformData>();

                entityXMLObj.readFromRC12XML(bodyElement, listPDatas, out foxproRC12Columns, out foxproRC12_SecondColumns, out foxproRC12_ThirdColumns);
                #endregion

                #region 儲存RC12 
                
                //儲存RC12資料表_五筆以下
                if (foxproRC12Columns.Count() > 0)
                {
                    if (_isWriteLog.ToUpper() == "Y")
                    {
                        string json = JsonConvert.SerializeObject(foxproRC12Columns);
                        _writeObj.writeToFile("新增DTata.Rc12資料: <<<" + json + ">>>");
                    }
                    entityXMLObj.AddToFoxpro("RC12", foxproRC12Columns, liAddRC12Columns, false);
                }
                    

                //儲存RC12資料表_超過5項處置
                if (foxproRC12_SecondColumns.Count() > 0)
                {
                    if (_isWriteLog.ToUpper() == "Y")
                    {
                        string json = JsonConvert.SerializeObject(foxproRC12_SecondColumns);
                        _writeObj.writeToFile("新增DTata.Rc12資料: <<<" + json + ">>>");
                    }
                    entityXMLObj.AddToFoxpro("RC12", foxproRC12_SecondColumns, liAddRC12Columns, false);
                }
                    

                //儲存RC12資料表_超過10項處置
                if (foxproRC12_ThirdColumns.Count() > 0)
                {
                    if (_isWriteLog.ToUpper() == "Y")
                    {
                        string json = JsonConvert.SerializeObject(foxproRC12_ThirdColumns);
                        _writeObj.writeToFile("新增DTata.Rc12資料: <<<" + json + ">>>");
                    }

                    entityXMLObj.AddToFoxpro("RC12", foxproRC12_ThirdColumns, liAddRC12Columns, false); 
                }                    
                #endregion
                #endregion
            }
        }

        private bool MoveFile(string fileName, string path1, string path2)
        {
            try
            {
                string newPath1 = Path.Combine(path1, fileName);
                string newPath2 = Path.Combine(path2, DateTime.Now.ToString("yyyyMMddhhmmss") + "_" + fileName);

                //搬移檔案到別的目錄
                if (File.Exists(newPath1))
                {
                    if (!Directory.Exists(path2))
                        Directory.CreateDirectory(path2);

                    File.Move(newPath1, newPath2);
                    if (_isWriteLog.ToUpper() == "Y")
                    {
                        _writeObj.writeToFile("搬移檔案成功: " + fileName );
                    }
                    return true;
                }
                else
                {
                    return false; //檔案不存在
                }
            }
            catch (Exception ex)
            {
                _writeObj.writeToFile(DateTime.Now.ToString("yyyyMMdd") + "_errorLog", _currentPath, "MoveFile:" + ex.Message.ToString());
                throw new Exception(ex.Message);
            }
        }
    }
}
