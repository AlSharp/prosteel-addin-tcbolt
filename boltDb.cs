using System;
using System.Data;
using System.Data.OleDb;
using System.Collections.Generic;

using Bentley.ProStructures.Miscellaneous;

namespace GSFBolt
{
public static class BoltDb
{
	public static string[] TablesToArray()
	{
		String connect = "Provider=Microsoft.ACE.OLEDB.12.0;data source=H:\\GSF\\WorkSpace\\ProStructures\\Data\\Bolts\\GSF_Bolts.mdb";
		OleDbConnection connection = new OleDbConnection(connect);
		connection.Open();
		DataTable tables = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[]{null,null,null,"TABLE"});
		int i = 1;
		string[] arr = new string[tables.Rows.Count + 1];
		arr[0] = "Type or select from the list";
		foreach(DataRow row in tables.Rows)
		{
			arr[i] = row[2].ToString();
			i++;
		}
		return arr;
	}
	public static string[] PopulateComboBoxWithDiameters(string Table)
	{
		List<double> list = new List<double>();
		String connect = "Provider=Microsoft.ACE.OLEDB.12.0;data source=H:\\GSF\\WorkSpace\\ProStructures\\Data\\Bolts\\GSF_Bolts.mdb";
		OleDbConnection connection = new OleDbConnection(connect);
		connection.Open();
		OleDbDataReader reader = null;
		OleDbCommand command = new OleDbCommand();
		command.Connection = connection;
		command.CommandText = "SELECT Diameter FROM [" + Table +"]";
		reader = command.ExecuteReader();
		PsUnits psUnits = new PsUnits();
		while (reader.Read())
		{
			double d = reader.GetDouble(0);
			list.Add(Math.Round(d, 3));
		}
		connection.Close();
		List<double> newList = new List<double>();
		foreach (double s in list)
		{
			if (!newList.Contains(s))
			{
				newList.Add(s);
			}
		}
		newList.Sort();
		double[] arr = newList.ToArray();
		string[] result = new string[arr.Length];
		for (int a = 0; a < arr.Length; a++)
		{
			result[a] = psUnits.ConvertToText(arr[a]);
		}
		return result;
	}
	public static string[] PopulateComboBoxWithLengths(string Table, double Criteria)
	{
		List<double> list = new List<double>();
		String connect = "Provider=Microsoft.ACE.OLEDB.12.0;data source=H:\\GSF\\WorkSpace\\ProStructures\\Data\\Bolts\\GSF_Bolts.mdb";
		OleDbConnection connection = new OleDbConnection(connect);
		connection.Open();
		OleDbDataReader reader = null;
		OleDbCommand command = new OleDbCommand();
		command.Connection = connection;
		command.CommandText = "SELECT Length FROM [" + Table +"] WHERE Diameter = " + Criteria;
		reader = command.ExecuteReader();
		PsUnits psUnits = new PsUnits();
		while (reader.Read())
		{
			double d = reader.GetDouble(0);
			list.Add(Math.Round(d, 3));
		}
		connection.Close();
		List<double> newList = new List<double>();
		foreach (double s in list)
		{
			if (!newList.Contains(s))
			{
				newList.Add(s);
			}
		}
		newList.Sort();
		double[] arr = newList.ToArray();
		string[] result = new string[arr.Length];
		for (int a = 0; a < arr.Length; a++)
		{
			result[a] = psUnits.ConvertToText(arr[a]);
		}
		return result;
	}
	public static double GetThreadedLength(string Table, double Diameter, double Length)
	{
		String connect = "Provider=Microsoft.ACE.OLEDB.12.0;data source=H:\\GSF\\WorkSpace\\ProStructures\\Data\\Bolts\\GSF_Bolts.mdb";
		OleDbConnection connection = new OleDbConnection(connect);
		OleDbCommand command = new OleDbCommand();
		command.Connection = connection;
		command.CommandText = "SELECT ThreadLength FROM [" + Table +"] WHERE Diameter = " + Diameter + " AND Length = " + Length;
		connection.Open();
		double result = (double)command.ExecuteScalar();
		connection.Close();
		return result;
	}
	public static double GetMinEmbedment(string Table, double Diameter, double Length)
	{
		String connect = "Provider=Microsoft.ACE.OLEDB.12.0;data source=H:\\GSF\\WorkSpace\\ProStructures\\Data\\Bolts\\GSF_Bolts.mdb";
		OleDbConnection connection = new OleDbConnection(connect);
		OleDbCommand command = new OleDbCommand();
		command.Connection = connection;
		command.CommandText = "SELECT MinEmbedment FROM [" + Table +"] WHERE Diameter = " + Diameter + " AND Length = " + Length;
		connection.Open();
		double result = (double)command.ExecuteScalar();
		connection.Close();
		return result;
	}
	public static string[] GetLinks(string Table, string Column)
	{
		List<string> list = new List<string>();
		String connect = "Provider=Microsoft.ACE.OLEDB.12.0;data source=H:\\GSF\\WorkSpace\\ProStructures\\Data\\Bolts\\GSF_Bolts.mdb";
		OleDbConnection connection = new OleDbConnection(connect);
		connection.Open();
		OleDbDataReader reader = null;
		OleDbCommand command = new OleDbCommand();
		command.Connection = connection;
		command.CommandText = "SELECT " + Column + " FROM [" + Table +"]";
		reader = command.ExecuteReader();
		while (reader.Read())
		{
			if(!reader.IsDBNull(0))
			{
			string s = reader.GetString(0);
			list.Add(s);
			}
		}
		connection.Close();
		string[] result = list.ToArray();
		return result;
	}
}

}