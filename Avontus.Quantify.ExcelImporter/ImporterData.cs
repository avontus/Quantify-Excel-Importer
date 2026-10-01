

using System;
using System.Data;
using System.Collections.Generic; 
using Avontus.Core;
using Avontus.Core.Data;
using Avontus.Rental.Library;  
using Avontus.Rental.Library.Utility;

namespace Avontus.Quantify.ExcelImporter
{
    [Serializable]
    public class ImporterData : CommandBase
    {
        private object _lookupData = null;  
        private string _importIdentifier; 

        public static object GetData(string importIdentifier)
        {
            var command = new ImporterData { };
            command._importIdentifier = importIdentifier; 

            DataPortal.Execute(command);
            return command._lookupData;
        }

        protected override void DataPortal_Execute()
        {
            switch (NormStr(_importIdentifier))
            {
                case "IMPORT:SCAFFOLDTAG":
                    Dictionary<string, Guid> priorityDict = new Dictionary<string, Guid>();
                    Dictionary<string, Guid> projectDict = new Dictionary<string, Guid>();
                    Dictionary<string, Guid> stepDict = new Dictionary<string, Guid>();
                    Dictionary<string, int> statusDict = new Dictionary<string, int>(); 

                    using (var ctx = DbUtil.GetConnectionManager())
                    using (var cmd = ctx.Connection.CreateCommand())
                    {
                        cmd.CommandText = "select * from Priority";
                        using (SafeDataReader dr = new SafeDataReader(cmd.ExecuteReader()))
                        {
                            while (dr.Read())
                            {
                                {
                                    string name = NormStrVar(dr.GetString("Name"));
                                    Guid Id = dr.GetGuid("PriorityID");
                                    priorityDict.Add(name, Id);
                                }
                            }
                        }

                        cmd.CommandText = "select * from Project";
                        using (SafeDataReader dr = new SafeDataReader(cmd.ExecuteReader()))
                        {
                            while (dr.Read())
                            {
                                {
                                    string name = NormStrVar(dr.GetString("Name"));
                                    Guid Id = dr.GetGuid("ProjectID");
                                    projectDict.Add(name, Id);
                                }
                            }
                        }

                        cmd.CommandText = "select * from ScaffoldTagStep";
                        using (SafeDataReader dr = new SafeDataReader(cmd.ExecuteReader()))
                        {
                            while (dr.Read())
                            {
                                {
                                    string name = NormStrVar(dr.GetString("Name"));
                                    Guid Id = dr.GetGuid("ScaffoldTagStepID");
                                    stepDict.Add(name, Id);
                                }
                            }
                        }


                        cmd.CommandText = "select * from ScaffoldTagStatus";
                        using (SafeDataReader dr = new SafeDataReader(cmd.ExecuteReader()))
                        {
                            while (dr.Read())
                            {
                                {
                                    int statusID = dr.GetInt16("ScaffoldTagStatusID"); 
                                    string name = dr.GetString("Name").ToUpper();
                                    statusDict.Add(name, statusID); 
                                }
                            }
                        }
                    }
                    _lookupData = new Tuple<Dictionary<string, Guid>, Dictionary<string, Guid>, Dictionary<string, Guid>,Dictionary<string,int>>
                        (priorityDict, projectDict, stepDict, statusDict);
                    break;
                case "IMPORT:SCAFFOLDTAGACTIVITY":
                    Dictionary<string, Guid> shiftDict = new Dictionary<string, Guid>();
                    Dictionary<Tuple<Guid,string>, Guid> typeDict = new Dictionary<Tuple<Guid, string>, Guid>();
                    Dictionary<Tuple<Guid,Int16, string>, Guid> jobListDict = new Dictionary<Tuple<Guid, Int16, string>, Guid>();
   
                    using (var ctx = DbUtil.GetConnectionManager())
                    using (var cmd = ctx.Connection.CreateCommand())
                    {
                        cmd.CommandText = "select * from Shift";
                        using (SafeDataReader dr = new SafeDataReader(cmd.ExecuteReader()))
                        {
                            while (dr.Read())
                            {
                                {
                                    string name = NormStrVar(dr.GetString("Name"));
                                    Guid Id = dr.GetGuid("ShiftID");
                                    shiftDict.Add(name, Id);
                                }
                            }
                        }

                        cmd.CommandText = "select * from ScaffoldTagActivityType";
                        using (SafeDataReader dr = new SafeDataReader(cmd.ExecuteReader()))
                        {
                            while (dr.Read())
                            {
                                {
                                    Guid locationID = dr.GetGuid("StockingLocationID");
                                    string name = dr.GetString("Name").ToUpper(); 
                                    Guid Id = dr.GetGuid("ScaffoldTagActivityTypeID");
                                    typeDict.Add(new Tuple<Guid,string>(locationID, name),Id);
                                }
                            }
                        }

                        cmd.CommandText = "select * from ScaffoldTagJobList";
                        using (SafeDataReader dr = new SafeDataReader(cmd.ExecuteReader()))
                        {
                            while (dr.Read())
                            {
                                {
                                    Guid locationID = dr.GetGuid("StockingLocationID");
                                    string name = dr.GetString("Name");
                                    Int16 listNumber = dr.GetInt16("ListNumber"); 
                                    Guid Id = dr.GetGuid("ScaffoldTagJobListItemID");
                                    jobListDict.Add(new Tuple<Guid, Int16, string>(locationID, listNumber, name), Id); 
                                }
                            }
                        }
                    }
                    _lookupData = new Tuple<Dictionary<string, Guid>, Dictionary<Tuple<Guid, string>, Guid>, Dictionary<Tuple<Guid, Int16, string>, Guid>>
                        (shiftDict, typeDict, jobListDict);
                    break; 
            }
        }

        public static string NormStr(string s)
        {
            return s.Replace(" ", "").ToUpper();
        }

        public static string NormStrVar(string s)
        {
            return s.ToUpper();
        }
    }
}