using NPOI.HSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Text.RegularExpressions;

namespace WaterReport

{    /// <summary>
     /// Copyright (C) 2016-2017 ZhaoLei 
     /// 报表数据处理基础类(数据基于SQLServer)
     /// </summary>
    class Service
    {
        public static Dictionary<int, string> varlist;
        public static Dictionary<int, string> waterlist;
        public static Dictionary<int, string> waterflowlist;

        #region  处理数据
        public static bool Exists(DateTime starttime, DateTime endtime,String when)
        {
            //读取配置信息
            string varsids = ConfigurationManager.AppSettings["varsid"];
            string waterids = ConfigurationManager.AppSettings["waterid"];
            string waterflows = ConfigurationManager.AppSettings["waterflow"];
            varlist = Ostring(varsids);
            waterlist = Ostring(waterids);
            waterflowlist = Ostring(waterflows);
            ArrayList vids = new ArrayList();
            ArrayList wids = new ArrayList();
            foreach (int _Key in varlist.Keys) vids.Add(_Key);
            foreach (int _Key in waterlist.Keys) wids.Add(_Key);
            //低于2012sqlserver 不支持
            // string sql = "SELECT id,varid,varValue,recordTime FROM (SELECT id,varid,varValue,recordTime,lag(varValue) over(ORDER BY id) AS prev_b FROM dbo.VarLogs WHERE recordTime>= '{0}' and recordTime<='{1}' AND varid ={2} GROUP By id, varid, varValue, recordTime  ) AS aa WHERE varValue<> prev_b OR prev_b IS NULL";
           // string sql = "SELECT distinct  A.id, A.varid,A.varValue,A.recordTime FROM (select row_number() over(order by recordTime ) as rowid,* FROM dbo.VarLogs  WHERE recordTime>= '{0}' and recordTime<='{1}' AND varid={2}) as A,(select row_number() over(order by recordTime) as rowid, *FROM dbo.VarLogs  WHERE recordTime >= '{0}' and recordTime <= '{1}' AND varid ={ 2}) as B WHERE A.rowid = B.rowid + 1 AND A.varValue<> b.varValue or(A.rowid= B.rowid and A.rowid= 1)";
            string sql = "SELECT distinct  A.id, A.varid,A.varValue,A.recordTime FROM (select row_number() over(order by recordTime ) as rowid,* FROM dbo.VarLogs  WHERE recordTime>= '{0}' and recordTime<='{1}' AND varid={2}) as A,(select row_number() over(order by recordTime) as rowid, * FROM dbo.VarLogs  WHERE recordTime >= '{3}' and recordTime <= '{4}' AND varid ={5}) as B WHERE A.rowid = B.rowid + 1 AND A.varValue<> b.varValue or (A.rowid= B.rowid and A.rowid= 1)";
            DataSet data = new DataSet();
            //每个水泵的数据
            foreach (int _Key in varlist.Keys) {

               string _sql = string.Format(sql, starttime, endtime, _Key, starttime, endtime, _Key);
                DataSet _data = SQLServerDAL.Query(_sql, _Key.ToString());
                if (_data.Tables[0].Rows.Count > 0)
                {
                    //处理第一条数据
                    if (_data.Tables[0].Rows[0]["varValue"].ToString() == "1")
                    {//周期以前打开是的数据
                        _data = setupdata(_data);
                    }
                    else
                    {
                        //上次打开是数据
                        // _data.Tables[0].Rows.RemoveAt(0);
                       _data = getupdata(_data);

                    }
                    //处理最后一条数据 
                    if (_data.Tables[0].Rows.Count != 0)
                    {
                        if (_data.Tables[0].Rows[_data.Tables[0].Rows.Count - 1].ToString() == "1")
                        {
                            DataRow dr = _data.Tables[0].NewRow();
                            dr["id"] = "9999999";
                            dr["varid"] = _Key;
                            dr["varValue"] = "0";
                            dr["recordTime"] = DBNull.Value;
                            _data.Tables[0].Rows.Add(dr);
                        }
                    }
                    else {
                        DataRow dr = _data.Tables[0].NewRow();
                        dr["id"] = "9999998";
                        dr["varid"] = _Key;
                        dr["varValue"] = "1";
                        dr["recordTime"] = DBNull.Value;
                        _data.Tables[_Key.ToString()].Rows.Add(dr);
                        DataRow _dr = _data.Tables[0].NewRow();
                        _dr["id"] = "9999999";
                        _dr["varid"] = _Key;
                        _dr["varValue"] = "0";
                        _dr["recordTime"] = DBNull.Value;
                        _data.Tables[0].Rows.Add(_dr);
                    }
                    
                }
                else {
                    DataRow dr = _data.Tables[0].NewRow();
                    dr["id"] = "9999998";
                    dr["varid"] = _Key;
                    dr["varValue"] = "1";
                    dr["recordTime"] = DBNull.Value;
                    _data.Tables[0].Rows.Add(dr);
                    DataRow _dr = _data.Tables[0].NewRow();
                    _dr["id"] = "9999999";
                    _dr["varid"] = _Key;
                    _dr["varValue"] = "0";
                    _dr["recordTime"] = DBNull.Value;
                    _data.Tables[0].Rows.Add(_dr);
                }
                data.Merge(_data, false, MissingSchemaAction.AddWithKey);
            }
            string sql_ = "SELECT TOP 1 log.id,log.recordTime,log.varId,log.varValue,log.varName FROM dbo.VarLogs AS log where log.varId={0}  AND log.recordTime<'{1}' ORDER BY log.recordTime DESC";

            foreach (DataTable dt in data.Tables)   //遍历所有的表的每一行 增加水位数据
            {
                foreach (int _Key in waterlist.Keys)
                {
                    dt.Columns.Add(_Key.ToString(), typeof(string));
                }
                foreach (DataRow mDr in dt.Rows)
                {
                    foreach (int _Key in waterlist.Keys)
                    {
                       DataSet _data = SQLServerDAL.Query(string.Format(sql_, _Key, mDr["recordTime"]), _Key.ToString());
                        if (_data.Tables[_Key.ToString()].Rows.Count > 0)
                        {
                            mDr[_Key.ToString()] = _data.Tables[_Key.ToString()].Rows[0]["varValue"];
                        }
                        else {
                            mDr[_Key.ToString()] = "-";
                        }
                      
                    }
                }
            }

            foreach (DataTable dt in data.Tables)   //遍历所有的datatable
            {
                foreach (DataRow mDr in dt.Rows)
                {
                    foreach (DataColumn mhr in dt.Columns)
                    {
                        Console.Write(mDr[mhr].ToString() + "   ");
                    }
                    Console.WriteLine();
                }
            }
            HSSFWorkbook Head = Excel_NPOI.CreateHead(data, varlist, waterlist, waterflowlist,  starttime,  endtime, when);
            return true;
        }
        #endregion

        #region  保存配置信息
        public static Dictionary<int, string> Ostring( string val) {
            Dictionary<int, string> var = new Dictionary<int, string>();
            string[] sArray = Regex.Split(val, ",", RegexOptions.IgnoreCase);
            foreach (string i in sArray) {
                string[] vArray = Regex.Split(i, "=", RegexOptions.IgnoreCase);
                var[Convert.ToInt32(vArray[0])] = vArray[1];
            } 
            return var;
        }
        #endregion


        #region 周期之前启动的一条数据

        public static DataSet setupdata(DataSet data )
        {
            string sql = "SELECT top 1 log.id,log.recordTime,log.varId,log.varValue,log.varName FROM dbo.VarLogs AS log where log.varId={0} AND log.varValue={1} AND log.recordTime<'{2}' ORDER BY log.recordTime DESC";
            sql = string.Format(sql, data.Tables[0].Rows[0]["varid"], 0, data.Tables[0].Rows[0]["recordTime"]);
            string _Key = data.Tables[0].TableName.ToString();
            //之前最后一次关闭的数据
            DataSet _data = SQLServerDAL.Query(string.Format(sql, _Key), _Key.ToString());
            DataRow dr = _data.Tables[0].NewRow();
            dr["id"] = _data.Tables[0].Rows.Count <= 0 ? "00000": _data.Tables[0].Rows[0]["id"].ToString();
            dr["varid"] =  data.Tables[0].Rows[0]["varid"];
            dr["varValue"] = _data.Tables[0].Rows.Count <= 0 ? "0" : _data.Tables[0].Rows[0]["varValue"].ToString();
            dr["recordTime"] = _data.Tables[0].Rows.Count <= 0 ? data.Tables[0].Rows[0]["recordTime"] : _data.Tables[0].Rows[0]["recordTime"];
            _data.Tables[0].Rows.InsertAt(dr, 0);
            //之前最后打开的一条数据
            sql = "SELECT top 1 log.id,log.recordTime,log.varId,log.varValue,log.varName FROM dbo.VarLogs AS log where log.varId={0} AND log.varValue={1} AND log.recordTime>'{2}' ORDER BY log.recordTime";
            sql = string.Format(sql, data.Tables[0].Rows[0]["varid"], 1, _data.Tables[0].Rows[0]["recordTime"]);
             _data = SQLServerDAL.Query(string.Format(sql, _Key), _Key.ToString());
            //Console.WriteLine(string.Format("{0}{1}",_data.Tables[_Key.ToString()].Rows[0]["recordTime"], data.Tables[_Key.ToString()].Rows[1]["recordTime"]));
            //替换周期内的启动数据
            if (_data.Tables[0].Rows[0]["recordTime"]!= data.Tables[0].Rows[0]["recordTime"] &&(_data.Tables[0].Rows[0]["varValue"]== data.Tables[0].Rows[0]["varValue"])) {
                Console.WriteLine("之前开启时");
                DataRow row = data.Tables[0].NewRow();
                row.ItemArray = _data.Tables[0].Rows[0].ItemArray;
                data.Tables[0].Rows.InsertAt(row, 0);
            }
            return data;
        }
        #endregion


        public static DataSet getupdata(DataSet data) {
            string sql = "SELECT top 1 id , varid,varValue,recordtime from VarLogs WHERE varid ={0} AND varValue={1} AND recordtime < '{2}' order by recordTime DESC";
            sql = string.Format(sql, data.Tables[0].Rows[0]["varid"], 1, data.Tables[0].Rows[0]["recordTime"]);
            string _Key = data.Tables[0].TableName.ToString();
            //上次开启的数据
            DataSet _data = SQLServerDAL.Query(string.Format(sql, _Key), _Key.ToString());
            if (_data.Tables[0].Rows.Count == 0)
            {//上次根本没有开启
                _data.Tables[0].Rows.RemoveAt(0);
            }
            else {
                DataRow row = data.Tables[0].NewRow();
                row.ItemArray = _data.Tables[0].Rows[0].ItemArray;
                data.Tables[0].Rows.InsertAt(row, 0);
            }
            return data;
        }
    }
}
