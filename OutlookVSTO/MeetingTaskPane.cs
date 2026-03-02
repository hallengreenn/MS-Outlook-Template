using System;
using System.Drawing;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMeetingAddin
{
    public partial class MeetingTaskPane : UserControl
    {
        private const string MEETING_TEMPLATE = @"
<p><strong>Formål med mødet</strong></p>
<p>Kort beskrivelse af, hvorfor vi mødes, og hvad vi skal opnå.</p>
<br>

<p><strong>Dagsorden/emner</strong></p>
<p>Liste over de punkter, der skal drøftes (fx status eller beslutninger)</p>
<br>

<p><strong>Roller og ansvar</strong></p>
<p>Hvem er mødeleder og hvem tager referat (hvis relevant).</p>
<br>

<p><strong>Beslutninger og næste skridt</strong></p>
<p>Afslut med at opsummere beslutninger og aftale opfølgning.</p>
<br>

<p><strong>Evt.</strong></p>
<p>Tid til spørgsmål eller andre punkter.</p>";

        public Outlook.AppointmentItem CurrentAppointment { get; set; }

        public MeetingTaskPane()
        {
            InitializeComponent();
        }

        private void btnInsertTemplate_Click(object sender, EventArgs e)
        {
            try
            {
                if (CurrentAppointment != null)
                {
                    // Indsæt skabelonen i appointment body
                    CurrentAppointment.Body = MEETING_TEMPLATE;

                    MessageBox.Show(
                        "Møde skabelon er blevet indsat!",
                        "Succes",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
                else
                {
                    MessageBox.Show(
                        "Ingen aktiv mødeaftale fundet.",
                        "Fejl",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Fejl ved indsættelse: {ex.Message}",
                    "Fejl",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }
    }
}
