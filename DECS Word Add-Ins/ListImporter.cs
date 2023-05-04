using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.LinkLabel;

namespace DecsWordAddIns
{
    internal class ListImporter
    {
        private const int MAX_LINES_PER_IMPORT = 1000;
        private const string PREAMBLE = "INSERT INTO MRN_LIST (MRN)\r\nVALUES(\r\n";
        private const string SUFFIX = ");";

        public ListImporter() 
        {
        }

        public void Scan(Document doc)
        {
            // Initialize the output .SQL file.
            string output_filename = Utilities.FormOutputFilename(filename: doc.FullName, filetype: ".sql");

            using (StreamWriter writer = new StreamWriter(output_filename))
            {
                writer.Write(PREAMBLE);
                int num_lines = doc.Paragraphs.Count;
                int lines_written = 0;
                int lines_written_this_chunk = 0;

                foreach (Paragraph para in doc.Paragraphs)
                {
                    if (para == null) continue;

                    string line = para.Range.Text.ToString().Trim();

                    if (line == null) continue;

                    string line_ending;

                    writer.Write("'" + line + "'");
                    lines_written++;
                    lines_written_this_chunk++;

                    if (lines_written < num_lines)
                    {
                        if (lines_written_this_chunk < MAX_LINES_PER_IMPORT)
                        {
                            line_ending = ",\r\n";
                        }
                        else
                        {
                            line_ending = SUFFIX + "\r\n" + PREAMBLE;
                            lines_written_this_chunk = 0;
                        }
                    }
                    else
                    {
                        line_ending = SUFFIX + "\r\n";
                    }
                    
                    writer.Write(line_ending);
                }
            }

            string message = "Created file '" + output_filename + "'.";
            MessageBox.Show(message, "Success");
        }
    }
}
