using ClassLibraryFoxDB;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using WriteEvent;

namespace entityXML
{
    /// <summary>
    /// XML和COOPER欄位對應的物件
    /// </summary>
    public class viewTransformData
    {
      
        public string elementAttribute
        {
            get;
            set;
        }

        public string foxproField
        {
            get;
            set;
        }

        public string value
        {
            get;
            set;
        }

        public OleDbType oledbType
        {
            get;
            set;
        }
        
    }

    
    public class entityXmlToFoxpro
    {

        private string _currentPath = Directory.GetCurrentDirectory();       

        private WrittingEventLog _writeObj = new WrittingEventLog();

        public entityXmlToFoxpro(string cooperPath)
        {
            foxproDB.CooperFolder = cooperPath;
        }

        public Dictionary<int, IList<viewTransformData>> ReadFromXML(string xmlfilePathName, string elementName, IList<viewTransformData> listViewTransformDataObjs)
        {
            //逐一取得viewTransformData 的欄位值
            try
            {
                XElement root = XElement.Load(xmlfilePathName);
                IEnumerable<XElement> xmlAttibutes = from xmldata in root.Elements(elementName) select xmldata;
                return GetXMLValue(listViewTransformDataObjs, xmlAttibutes);
            }
            catch (Exception ex)
            {
                _writeObj.writeToFile(DateTime.Now.ToString("yyyyMMdd") + "_errorLog", _currentPath, "ReadFromXML:71" +  ex.Message.ToString());
                throw new Exception(ex.Message);
            }          
        }

        private static string customerFormatValue(string strFormat, string value)
        {
            try
            {
                string[] strTemp = strFormat.ToUpper().Split(':');
                int defNumber = 0;
                int.TryParse(strTemp[1], out defNumber);
                switch (strTemp[0])
                    {
                        case "ZERO"://左邊要補零                            
                            return value.PadLeft(defNumber, '0');
                        case "SPACE": //左邊補上空白                            
                            return value.PadLeft(defNumber, ' ');
                        case "*": //乘上某個數字
                            int  multiplication =0;
                            string tempValue = (value.Contains(".")) ? value.Substring(0, value.LastIndexOf(".")): value ;
                            int.TryParse(tempValue, out multiplication);
                            return (multiplication * defNumber).ToString();
                    default:
                            return value;
                    }
            }
            catch (Exception ex)
            {
                throw new Exception("customerFormatValue:" + ex.Message);
            }
        }

        public static string definitionValue(string defValue, string xmlValue)
        {
            try
            {
                if (defValue == null)
                    return xmlValue;

                string[] strValueFormat = defValue.Split(',');
                int deCount = strValueFormat.Count();
                while (deCount > 0)
                {
                    xmlValue = customerFormatValue(strValueFormat[strValueFormat.Count() - deCount], xmlValue);
                    deCount--;
                }
                return xmlValue;
            }
            catch (Exception ex)
            {
                throw new Exception("definitionValue:" + ex.Message);
            }
            
        }

        public  Dictionary<int, IList<viewTransformData>> GetXMLValue(IList<viewTransformData> listViewTransformDataObjs, IEnumerable<XElement> xmlAttibutes)
        {
            Dictionary<int, IList<viewTransformData>> dataXMLRows = new Dictionary<int, IList<viewTransformData>>();
            string strValue = "";
            int dataRowIndex = 0;
            foreach (XElement xmlAttr in xmlAttibutes)
            {
                IList<viewTransformData> readElements = new List<viewTransformData>();

                foreach (viewTransformData viewTransFormDataObj in listViewTransformDataObjs)
                {
                    string strElement = (string)xmlAttr.Element(viewTransFormDataObj.elementAttribute);
                    if (strElement == null)
                    {
                        continue;
                    }

                    if ((viewTransFormDataObj.oledbType == OleDbType.Decimal || viewTransFormDataObj.oledbType == OleDbType.Double || viewTransFormDataObj.oledbType == OleDbType.SmallInt || viewTransFormDataObj.oledbType == OleDbType.Numeric) && strElement.Trim() == "")
                    {
                        strValue = "0";
                    }
                    else
                    {
                        strValue = (string)xmlAttr.Element(viewTransFormDataObj.elementAttribute);
                        
                        if (viewTransFormDataObj.value != null)
                        {
                            //有些欄位值有客製化需求，例如左邊補零，左邊補空白等
                            strValue = definitionValue(viewTransFormDataObj.value, strValue);
                        }                        
                    }
                    readElements.Add(new viewTransformData()
                    {
                        elementAttribute = viewTransFormDataObj.elementAttribute,
                        foxproField = viewTransFormDataObj.foxproField,
                        oledbType = viewTransFormDataObj.oledbType,
                        value = strValue
                    });
                }

                dataXMLRows.Add(dataRowIndex, readElements);
                dataRowIndex++;
            }
            return dataXMLRows;
        }
             

        public IList<viewTransformData> readFromXML(IEnumerable<XElement> xmlAttibutes, IList<viewTransformData> listViewTransformDataObjs)
        {
            IList<viewTransformData> removeColumns = new List<viewTransformData>();            
            try
            {
                foreach (viewTransformData viewData in listViewTransformDataObjs)
                {
                    var element = xmlAttibutes.Elements(viewData.elementAttribute);
                    if (element.Count() == 0)
                    {
                        removeColumns.Add(viewData);
                        continue;
                    }
                    viewData.value = definitionValue(viewData.value, element.FirstOrDefault().Value.ToString());
                }

                foreach (viewTransformData remove in removeColumns)
                {
                    listViewTransformDataObjs.Remove(remove);
                }


                return listViewTransformDataObjs;
            }
            catch (Exception ex)
            {
                throw new Exception("readFromXML:199 "+ ex.Message);
            }
        }
        
        public void readFromRC12XML(IEnumerable<XElement> xmlAttributes, IList<viewTransformData> listViewTransformDataobjs, out IList<viewTransformData> foxproRC12Columns, out IList<viewTransformData> foxproRC12_SecondColumns, out IList<viewTransformData>  foxproRC12_ThirdColumns)
        {
            try
            {
                int count = 1;
                foxproRC12Columns = new List<viewTransformData>();
                foxproRC12_SecondColumns = new List<viewTransformData>();
                foxproRC12_ThirdColumns = new List<viewTransformData>();

                foreach (var pDataElement in xmlAttributes.Elements("pdata"))
                {

                    foreach (viewTransformData viewTransformDataObj in listViewTransformDataobjs)
                    {
                        var element = pDataElement.Elements(viewTransformDataObj.elementAttribute);
                        if (element.Count() == 0)
                            continue;

                        if (count <= 5)
                        {
                            foxproRC12Columns.Add(new viewTransformData()
                            {
                                elementAttribute=viewTransformDataObj.elementAttribute,
                                foxproField = viewTransformDataObj.foxproField + count.ToString(),
                                oledbType = OleDbType.Char,
                                value = definitionValue(viewTransformDataObj.value, element.FirstOrDefault().Value.ToString())
                            });
                        }
                        else if (count > 5 && count <= 10)
                        {
                            foxproRC12_SecondColumns.Add(new viewTransformData()
                            {
                                foxproField = viewTransformDataObj.foxproField + (count - 5),
                                oledbType = OleDbType.Char,
                                value = definitionValue(viewTransformDataObj.value, element.FirstOrDefault().Value.ToString())
                            });
                        }
                        else if (count > 10 && count > 15)
                        {
                            foxproRC12_ThirdColumns.Add(new viewTransformData()
                            {
                                foxproField = viewTransformDataObj.foxproField + (count - 10),
                                oledbType = OleDbType.Char,
                                value = definitionValue(viewTransformDataObj.value, element.FirstOrDefault().Value.ToString())
                            });
                        }
                    }
                    count++;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("readFromRC12XML:" + ex.Message);
            }

        }

        public bool AddToFoxpro(string tableName, IList<viewTransformData> listViewTransformDataObjs, List<columnsData>additionColumns =null, bool haveIkey=true)
        {
            try
            {
                List<columnsData> liColumns = new List<columnsData>();

                /* 自動產生其他欄位的預設值 */
                listViewTransformDataObjs = autoGenerateDefault(tableName, listViewTransformDataObjs,haveIkey);

                foreach (viewTransformData viewTransformDataObj in listViewTransformDataObjs)
                {
                    //當欄位空白時，直接略過
                    if (viewTransformDataObj.foxproField == "")
                        continue;

                    liColumns.Add(new columnsData()
                    {
                         strFileName=viewTransformDataObj.foxproField,
                         strValue = viewTransformDataObj.value,
                         oledbTypeValue = viewTransformDataObj.oledbType
                    });
                }

                //額外增加自訂的欄位
                if (additionColumns != null)
                {
                    liColumns.AddRange(additionColumns.ToArray());
                }

                foxproDB.addWithParameter(tableName, liColumns);
                return true;
            }
            catch (Exception ex)
            {
                _writeObj.writeToFile(DateTime.Now.ToString("yyyyMMdd") + "_errorLog", _currentPath, "AddToFoxpro:" + ex.Message.ToString());
                throw new Exception("AddToFoxpro" + ex.Message);
            }
        }

        public bool EditToFoxpro(string tableName, IList<viewTransformData> listViewTransformDataObjs, int ikey)
        {
            try
            {
                List<columnsData> liColumns = new List<columnsData>();

                foreach (viewTransformData viewTransformDataObj in listViewTransformDataObjs)
                {
                    liColumns.Add(new columnsData()
                    {
                        strFileName = viewTransformDataObj.foxproField,
                        strValue = viewTransformDataObj.value,
                        oledbTypeValue = viewTransformDataObj.oledbType
                    });
                }

                liColumns.Add(new columnsData()
                {
                     strFileName="ikey",
                     strValue=ikey.ToString(),
                     oledbTypeValue= OleDbType.SmallInt
                });

                liColumns.Add(new columnsData()
                {
                    strFileName = "異動方式14",
                    strValue = "M",
                    oledbTypeValue = OleDbType.Char
                });
                foxproDB.updateWithParameter(tableName,liColumns);
                return true;
            }
            catch (Exception ex)
            {
                _writeObj.writeToFile(DateTime.Now.ToString("yyyyMMdd") + "_errorLog", _currentPath, "EditToFoxpro:" + ex.Message.ToString());
                throw new Exception("EditToFoxpro" + ex.Message);
            }
        }
        
        public bool DeleteWithZapToFoxpro(string tableName)
        {
            try
            {
                foxproDB.zapDataTable(tableName);
                return true;
            }
            catch (Exception ex)
            {
                throw new Exception("DeleteWithZapToFoxpro:" + ex.Message);
            }
        }

        public bool CheckIdentityExist(string tableName, string checkExistColumnName,  string value, out int ikey)
        {
            try
            {
                //object checkDataExist = foxproDB.selectQueryWithExecuteScalar("select " + checkExistColumnName + " from " + tableName + " where " + checkExistColumnName + " ='"+value +"'");

                DataTable dtCheckDataExist = foxproDB.selectQueryWithDataTable("select " + checkExistColumnName + ",ikey from " + tableName + " where " + checkExistColumnName + " ='"+value +"'");

                if (dtCheckDataExist.Rows.Count > 0)
                {
                    int.TryParse(dtCheckDataExist.Rows[0]["ikey"].ToString(), out ikey);
                    return true;
                }
                else
                {
                    ikey = 0;
                    return false;
                }
                //if (checkDataExist == null)
                //    return false;
                //else
                //    return true;
            }
            catch (Exception ex)
            {
                _writeObj.writeToFile(DateTime.Now.ToString("yyyyMMdd") + "_errorLog", _currentPath, "CheckIdentityExist:" + ex.Message.ToString());
                throw new Exception("CheckIdentityExist:" + ex.Message);
            }
        }

        /// <summary>
        /// 自動取得流水號
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="flowColumnName"></param>
        /// <returns></returns>
        public string getNewDoctorID(string tableName, string flowColumnName)
        {
            try
            {
                object doctor = foxproDB.selectQueryWithExecuteScalar(" select max(" + flowColumnName + ")  from " + tableName );

                if (doctor == null)
                    return "001";
                else
                {
                    int intDoctorId = 0;
                    int.TryParse(doctor.ToString(), out intDoctorId);
                    return (intDoctorId + 1).ToString().PadLeft(3, '0');
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        

        public bool CheckDoctorIdentityExist(string tableName, string checkExistColumnName, string value, out int ikey)
        {
            try
            {
                string strChiDate = (DateTime.Now.Year - 1911).ToString() + DateTime.Now.ToString("MMdd");
                DataTable dtDoctors = foxproDB.selectQueryWithDataTable(" select " + checkExistColumnName + ",ikey  from " + tableName + " where " + checkExistColumnName + " ='" + value + "' and ( 離職日期=' '   or 離職日期>'" + strChiDate + "' ) ");

                if (dtDoctors.Rows.Count > 0)
                {
                    int.TryParse(dtDoctors.Rows[0]["ikey"].ToString(), out ikey);
                    return true;
                }
                else
                {
                    ikey = 0;
                    return false;
                }

                //object doctor = foxproDB.selectQueryWithExecuteScalar(" select " + checkExistColumnName + "  from " + tableName + " where " + checkExistColumnName + " ='" + value + "' and ( 離職日期=' '   or 離職日期>'" + strChiDate + "' ) ");

                //if (doctor == null)
                //    return false;
                //else
                //    return true;
            }
            catch (Exception ex)
            {
                _writeObj.writeToFile(DateTime.Now.ToString("yyyyMMdd") + "_errorLog", _currentPath, "CheckIdentityExist:" + ex.Message.ToString());
                throw new Exception("CheckIdentityExist:" + ex.Message);
            }
        }
        
        public IList<viewTransformData> autoGenerateDefault(string tableName, IList<viewTransformData> liviewTransFormData,bool haveIkey=true)
        {
            try
            {
                DataTable dt;

                if (haveIkey == true)
                {
                    dt = foxproDB.selectQueryWithDataTable("SELECT  TOP 1 * FROM " + tableName + " ORDER BY IKEY DESC");
                }
                else
                {
                    dt = foxproDB.selectQueryWithDataTable("SELECT * FROM " + tableName);
                }
                
                foreach (DataColumn column in dt.Columns)
                {
                    var defaultField = (from q in liviewTransFormData
                                        where q.foxproField == column.ColumnName
                                        select q).FirstOrDefault();

                    if (defaultField == null)
                    {
                        viewTransformData newFiled = new viewTransformData();
                        if (column.ColumnName.ToUpper() == "IKEY")
                        {
                            int intNewIkey=0;
                            int.TryParse(dt.Rows[0][column.ColumnName].ToString(), out intNewIkey);
                            newFiled.value = (intNewIkey+1).ToString();
                            newFiled.oledbType = OleDbType.SmallInt;
                            newFiled.foxproField = column.ColumnName;
                        }
                        else
                        {
                            newFiled = _convertOdbcType(column.ColumnName, column.DataType.Name);
                        }
                        
                        liviewTransFormData.Add(newFiled);
                    }                    
                }
                return liviewTransFormData;
            }
            catch (Exception ex)
            {
                _writeObj.writeToFile(DateTime.Now.ToString("yyyyMMdd") + "_errorLog", _currentPath, "autoGenerateDefault:" + ex.Message.ToString());
                throw new Exception("autoGenerateDefault:" + ex.Message);
            }
        }

        protected viewTransformData _convertOdbcType(string strFileName, string strDataType)
        {
            try
            {
                viewTransformData defaultValue = new viewTransformData();
                switch (strDataType)
                {
                    case "String":
                        defaultValue.oledbType = OleDbType.Char;
                        defaultValue.foxproField = strFileName;
                        defaultValue.value="";
                        return defaultValue;
                    case "Int32":
                        defaultValue.oledbType = OleDbType.SmallInt;
                        defaultValue.foxproField = strFileName;
                        defaultValue.value = "0";
                        return defaultValue;
                    case "Boolean":
                        defaultValue.oledbType = OleDbType.Boolean;
                        defaultValue.foxproField = strFileName;
                        defaultValue.value = "False";
                        return defaultValue;
                    case "Decimal":
                        defaultValue.oledbType = OleDbType.Double;
                        defaultValue.foxproField = strFileName;
                        defaultValue.value = "0";
                        return defaultValue;
                    case "Date":
                        defaultValue.oledbType = OleDbType.Date;
                        defaultValue.foxproField = strFileName;
                        defaultValue.value = "1999/01/01";
                        return defaultValue;
                    case "DateTime":
                        defaultValue.oledbType = OleDbType.DBTimeStamp;
                        defaultValue.foxproField = strFileName;
                        defaultValue.value = "0";
                        return defaultValue;
                    case "Double":
                        defaultValue.oledbType = OleDbType.Double;
                        defaultValue.foxproField = strFileName;
                        defaultValue.value = "0";
                        return defaultValue;
                    default:
                        defaultValue.oledbType = OleDbType.VarChar;
                        defaultValue.foxproField = strFileName;
                        defaultValue.value = "";
                        return defaultValue;
                }
            }
            catch (Exception ex)
            {
                _writeObj.writeToFile(DateTime.Now.ToString("yyyyMMdd") + "_errorLog", _currentPath, "_convertOdbcType:" + ex.Message.ToString());
                throw new Exception("_convertOdbcType" + ex.Message);
            }
        }
    }
}
