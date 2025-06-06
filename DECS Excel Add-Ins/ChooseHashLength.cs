using System;
using System.IO;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace DECS_Excel_Add_Ins
{
    public partial class ChooseHashLength : Form
    {
        private string fileName = "hash_length.xml";
        private string filePath = string.Empty;
        public int hashLength;

        public ChooseHashLength()
        {
            InitializeComponent();
            DefineXmlFilename();
            ReadConfigFile();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            hashLength = (int)hashLengthUpDown.Value;
            WriteConfigFile(hashLength);
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void DefineXmlFilename()
        {
            filePath = Path.Combine(Path.GetTempPath(), fileName);
        }

        private void ReadConfigFile()
        {
            try
            {
                using (var reader = new StreamReader(filePath))
                {
                    // Read hash length value used last time.
                    HashLength hashLen = (HashLength)new XmlSerializer(typeof(HashLength)).Deserialize(reader);

                    // Initialize updown control.
                    hashLengthUpDown.Value = hashLen.hashLength;
                }
            }
            catch (DirectoryNotFoundException) { }
            catch (FileNotFoundException) { }
            catch (UnauthorizedAccessException) { }
        }

        private void WriteConfigFile(int hashLength)
        {
            // Ensure directory exists.
            Directory.CreateDirectory(Path.GetDirectoryName(filePath));

            using (var writer = new System.IO.StreamWriter(filePath))
            {
                var serializer = new XmlSerializer(typeof(HashLength));
                HashLength hashLen = new HashLength(hashLength);
                serializer.Serialize(writer, hashLen);
                writer.Flush();
            }
        }
    }
    public class HashLength
    {
        public int hashLength;

        public HashLength() { }

        public HashLength(int hashLength) { this.hashLength = hashLength; }
    }
}
