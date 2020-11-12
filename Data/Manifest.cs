using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlProto.Data
{
    public class Manifest
    {
        public string Title { get; set; }

        public IEnumerable<Plane> Planes { get; set; }

        public bool GenerateDocument(string path)
        {
            try
            {
                var spreadsheet = new ExcelDocument(Path.Join(path, $"{Title}.xlsx"));

                Planes.ToList().ForEach(plane => GenerateSheet(plane, spreadsheet));

                spreadsheet.SaveAndClose();

                return true;
            }
            catch
            {
                Console.WriteLine("You really fucked up");
                return false;
            }
        }

        static void GenerateSheet(Plane plane, ExcelDocument spreadsheet)
        {
            var worksheet = spreadsheet.CreateWorksheet(spreadsheet.Workbook);

            var sheet = spreadsheet.CreateSheet(worksheet, plane.Name);
            var sheetData = worksheet.Worksheet.GetFirstChild<SheetData>();

            sheetData.Append(GenerateHeader());

            var people = plane.People.ToList();

            people.ForEach(person =>
            {
                var index = (uint)people.IndexOf(person) + 2;
                sheetData.Append(GenerateRow(person, index));
            });

            ExcelDocument.ConfigureAutoFilter(worksheet, "A1:I1");

        }

        static Row GenerateHeader()
        {
            // FOR SOME STUPID FUCKING REASON, YOU CANNOT APPLY STYLES TO ROWS
            var row = new Row() { RowIndex = 1 };
            
            row.InsertAt(GenerateCell("A1", "SSN", true), 0);
            row.InsertAt(GenerateCell("B1", "Rank", true), 1);
            row.InsertAt(GenerateCell("C1", "Last Name", true), 2);
            row.InsertAt(GenerateCell("D1", "First Name", true), 3);
            row.InsertAt(GenerateCell("E1", "Middle Name", true), 4);
            row.InsertAt(GenerateCell("F1", "Title", true), 5);
            row.InsertAt(GenerateCell("G1", "Organization", true), 6);
            row.InsertAt(GenerateCell("H1", "DoD ID", true), 7);
            row.InsertAt(GenerateCell("I1", "Occupation", true), 8);

            return row;
        }

        static Row GenerateRow(Person person, uint index)
        {
            var row = new Row() { RowIndex = UInt32Value.FromUInt32(index) };

            row.InsertAt(GenerateCell($"A{index}", person.Ssn), 0);
            row.InsertAt(GenerateCell($"B{index}", person.Rank), 1);
            row.InsertAt(GenerateCell($"C{index}", person.LastName), 2);
            row.InsertAt(GenerateCell($"D{index}", person.FirstName), 3);
            row.InsertAt(GenerateCell($"E{index}", person.MiddleName), 4);
            row.InsertAt(GenerateCell($"F{index}", person.Title), 5);
            row.InsertAt(GenerateCell($"G{index}", person.Organization), 6);
            row.InsertAt(GenerateCell($"H{index}", person.DodId), 7);
            row.InsertAt(GenerateCell($"I{index}", person.Occupation), 8);

            return row;
        }

        static Cell GenerateCell(string cellRef, string value, bool isHeader = false, CellValues type = CellValues.String) => new Cell()
        {
            CellReference = cellRef,
            CellValue = new CellValue(value),
            DataType = new EnumValue<CellValues>(type),
            // STYLES MUST BE APPLIED TO CELLS
            StyleIndex = isHeader ? 1U : 0U
        };

        public static Manifest GenerateManifest(string title) => new Manifest
        {
            Title = title,
            Planes = new List<Plane>
            {
                new Plane
                {
                    Name = "Cessna 208 B Grand Caravan",
                    People = new List<Person>
                    {
                        new Person
                        {
                            Rank = "CTR",
                            LastName = "Van Damme",
                            FirstName = "John",
                            MiddleName = "Claude",
                            DodId = "1234567891",
                            Occupation = "Engineer",
                            Title = "Lead Engineer",
                            Organization = "Arma",
                            Ssn = "111-11-1111"
                        },
                        new Person
                        {
                            Rank = "CTR",
                            LastName = "Norris",
                            FirstName = "Chuck",
                            MiddleName = "Goddamn",
                            DodId = "1234567892",
                            Occupation = "Engineer",
                            Title = "Engineer",
                            Organization = "Arma",
                            Ssn = "222-22-2222"
                        }
                    }
                },
                new Plane
                {
                    Name = "Daher TBM 930",
                    People = new List<Person>
                    {
                        new Person
                        {
                            Rank = "CTR",
                            LastName = "Stegall",
                            FirstName = "Steven",
                            MiddleName = "Diabetes",
                            DodId = "1234567893",
                            Occupation = "Engineer",
                            Title = "Senior Engineer",
                            Organization = "Arma",
                            Ssn = "333-33-3333"
                        },
                        new Person
                        {
                            Rank = "CTR",
                            LastName = "Lee",
                            FirstName = "Bruce",
                            DodId = "1234567894",
                            Occupation = "Project Manager",
                            Title = "Site Lead",
                            Organization = "Arma",
                            Ssn = "444-44-4444"
                        }
                    }
                }
            }
        };
    }
}