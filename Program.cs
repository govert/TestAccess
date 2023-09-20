using System;
using Microsoft.Office.Interop.Access.Dao;

class Program
{
    static void Main()
    {
        string dbPath = $@"C:\temp\sample-{DateTime.Now:yyyyMMdd-HHmmss}.mdb";

        try
        {
            // Create a new database
            var engine = new DBEngine();
            Database db = engine.CreateDatabase(dbPath, /* LanguageConstants.dbLangGeneral */ ";LANGID=0x0409;CP=1252;COUNTRY=0");

            Console.WriteLine("Database created successfully!");

            // Create a table with Id and Name fields
            TableDef table = db.CreateTableDef("Persons");
            Field idField = table.CreateField("Id", DataTypeEnum.dbLong);
            idField.Attributes = (int)FieldAttributeEnum.dbAutoIncrField;
            table.Fields.Append(idField);
            table.Fields.Append(table.CreateField("Name", DataTypeEnum.dbText, 255));

            db.TableDefs.Append(table);
            db.Close();

            Console.WriteLine("Table created successfully!");

            // Insert some generated data into the database
            db = engine.OpenDatabase(dbPath);
            Recordset rs = db.OpenRecordset("Persons", RecordsetTypeEnum.dbOpenDynaset);

            for (int i = 0; i < 5; i++)
            {
                rs.AddNew();
                rs.Fields["Name"].Value = "Name " + i;
                rs.Update();
            }
            rs.Close();

            Console.WriteLine("Data inserted successfully!");

            // Perform a SELECT query and write results to console
            rs = db.OpenRecordset("SELECT * FROM Persons", RecordsetTypeEnum.dbOpenSnapshot);
            while (!rs.EOF)
            {
                Console.WriteLine($"Id: {rs.Fields["Id"].Value}, Name: {rs.Fields["Name"].Value}");
                rs.MoveNext();
            }

            rs.Close();
            db.Close();
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}


