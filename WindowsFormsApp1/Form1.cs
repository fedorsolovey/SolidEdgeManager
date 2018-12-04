using System;
using System.Windows.Forms;
using System.IO;
using SolidEdgeFramework;
using System.Runtime.InteropServices;
using SolidEdge.Part.Variables;
using System.Drawing;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private enum ColumnType { name, value };

        private double step = 0.01;

        private int selectedRow = 0;

        SolidEdgeFramework.Application application = null;
        SolidEdgeFramework.Documents documents = null;
        SolidEdgeFramework.SolidEdgeDocument document = null;
        SolidEdgeFramework.Variables variables = null;
        VariableList variableList = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.ColumnCount = 2;
            dataGridView1.Columns[(int)ColumnType.name].Name = @"Имя";
            dataGridView1.Columns[(int)ColumnType.value].Name = @"Значение";
            dataGridView1.Columns[(int)ColumnType.name].ReadOnly = true;
            dataGridView1.Columns[(int)ColumnType.value].ReadOnly = false;
            dataGridView1.Columns[(int)ColumnType.name].ValueType = typeof(string);
            dataGridView1.Columns[(int)ColumnType.value].ValueType = typeof(double);
            dataGridView1.Columns[(int)ColumnType.value].DefaultCellStyle.Format = "N3";

            numeric.Value = (decimal)step;

            ConnectToSolid();
            UpdateVariables();
        }

        private void UpdateVariables()
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Refresh();

            variables = (Variables)document.Variables;

            variableList = (VariableList)variables
            .Query(pFindCriterium: "*", 
            NamedBy: SolidEdgeConstants.VariableNameBy.seVariableNameByBoth,
            VarType: SolidEdgeConstants.VariableVarType.SeVariableVarTypeBoth);

            dynamic variableListItem = null;

            for (int i = 1; i <= variableList.Count; i++)
            {
                variableListItem = variableList.Item(i);
                dynamic f = variableListItem.Formula.Length == 0;
                string value = variableListItem.Value.ToString("0.000"); ;
                object[] row = { variableListItem.DisplayName, variableListItem.Value };
                dataGridView1.Rows.Add(row);
                dataGridView1.Rows[i - 1].DefaultCellStyle.BackColor = (variableListItem.Formula.Length == 0 ? Color.White : Color.Coral);
            }
        }

        private void DataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            double newValue = (double)dataGridView1.Rows[e.RowIndex].Cells[(int)ColumnType.value].Value;
            dynamic variableListItem = variableList.Item(selectedRow + 1);
            variableListItem.Value = newValue;
        }

        private void DataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            selectedRow = dataGridView1.CurrentCell.RowIndex;
        }

        private void UpButton_Click(object sender, EventArgs e)
        {
            dynamic variableListItem = variableList.Item(selectedRow + 1);
            UpdateValue(variableListItem.Value + step);
        }

        private void DownButton_Click(object sender, EventArgs e)
        {
            dynamic variableListItem = variableList.Item(selectedRow + 1);
            if (variableListItem.Value - step >= 0)
            {
                UpdateValue(variableListItem.Value - step);
            }
        }

        private void UpdateValue(double newValue)
        {
            DataGridViewCell cell = dataGridView1.Rows[selectedRow].Cells[(int)ColumnType.value];
            cell.Value = newValue;
        }

        private void NumericValueChanged(object sender, EventArgs e)
        {
            step = decimal.ToDouble(((NumericUpDown)sender).Value);
        }

        private void ConnectToSolid()
        {
            try
            {
                OleMessageFilter.Register();
                application = ConnectToSolidEdge(true);
                application.Visible = true;
                application.Activate();

                // Get a reference to the Documents collection.
                documents = application.Documents;

                application.Activate();
                // This check is necessary because application.ActiveDocument will throw an
                // exception if no documents are open...
                if (documents.Count > 0)
                {
                    // Attempt to connect to ActiveDocument.
                    document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;
                }

                // Make sure we have a document.
                if (document == null)
                {
                    throw new Exception("No active document.");
                }

                variables = (Variables)document.Variables;
            }
            catch (Exception ex)
            {

            }
            finally
            {
                OleMessageFilter.Revoke();
            }
        }

        /// <summary>
        /// Connects to a running instance of Solid Edge with an option to start if not running.
        /// </summary>
        public static SolidEdgeFramework.Application ConnectToSolidEdge(bool startIfNotRunning)
        {
            try
            {
                // Attempt to connect to a running instance of Solid Edge.
                return (SolidEdgeFramework.Application)
                    Marshal.GetActiveObject("SolidEdge.Application");
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                // Failed to connect.
                if (ex.ErrorCode == -2147221021 /* MK_E_UNAVAILABLE */)
                {
                    if (startIfNotRunning)
                    {
                        // Start Solid Edge.
                        return (SolidEdgeFramework.Application)
                            Activator.CreateInstance(Type.GetTypeFromProgID("SolidEdge.Application"));
                    }
                    else
                    {
                        throw new System.Exception("Solid Edge is not running.");
                    }
                }
                else
                {
                    throw;
                }
            }
            catch
            {
                throw;
            }
        }
    }
}
