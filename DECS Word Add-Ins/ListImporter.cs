using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.LinkLabel;

namespace DecsWordAddIns
{
    /**
     * @brief Reads a list from this Word document & outputs as a .SQL file.
     */ 
    internal class ListImporter
    {
        private readonly string[] IGNORED_WORDS = { "MRN" };
        private const int MAX_LINES_PER_IMPORT = 1000;
        private const string PREAMBLE = "USE [REL_CLARITY];\r\n\r\n";
        private const string SEGMENT_START = "INSERT INTO #MRN_LIST (MRN)\r\nVALUES\r\n";

        public ListImporter() { }

        /// <summary>
        /// Main method: Scans the document and builds the SQL file to import the list.
        /// </summary>
        /// <param name="doc">Word @c Document object</param>
        public void Scan(Document doc)
        {
            // Initialize the output .SQL file.
            (StreamWriter writer, string outputFilename) = Utilities.OpenOutput(
                input_filename: doc.FullName,
                filetype: ".sql"
            );

            writer.Write(PREAMBLE + SEGMENT_START);
            int num_lines = doc.Paragraphs.Count;
            int lines_written = 0;
            int lines_written_this_chunk = 0;

            foreach (Paragraph para in doc.Paragraphs)
            {
                if (para == null)
                    continue;

                string line = para.Range.Text.ToString().Trim();

                if (line == null)
                    continue;

                // If the line is just "MRN", ignore it.
                if (IGNORED_WORDS.Contains(line))
                    continue;

                string line_ending;

                writer.Write("('" + line + "')");
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
                        line_ending = ";\r\n\r\n" + SEGMENT_START;
                        lines_written_this_chunk = 0;
                    }
                }
                else
                {
                    line_ending = ";\r\n";
                }

                writer.Write(line_ending);
            }

            writer.Close();
            Process.Start(outputFilename);
        }
    }
}
