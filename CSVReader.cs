//Reads from a CSV file (or other comma delimited files

/// <summary>
        /// Reads from CSV files and adds values to a DataTable
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static DataTable ReadCSV(string fileName)
        {
            //Creates new table
            DataTable table = new DataTable();

            //columns of fincen csv file
            for (int i = 0; i < 18; i++)
            {
                table.Columns.Add();
            }

            //Starts reading the file
            using (var reader = new StreamReader(fileName))
            {
                //keeps looping until the stream stops seeing data
                while (!reader.EndOfStream)
                {
                    //string array to store the row of data
                    var values = reader.ReadLine().Split(',');
                    //new row to store the data
                    DataRow row = table.NewRow();
                    
                    //iterate over all columns to fill the row
                    for (int j = 0; j < 18; j++)
                    {
                        //add the array to the DataRow
                        row[j] = values[j];
                    }

                    //Add row to the table
                    table.Rows.Add(row);
                }
            }

            //Return the table
            return table;
        }
