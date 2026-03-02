using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools;

namespace OutlookMeetingAddin
{
    public partial class ThisAddIn
    {
        // Dictionary til at holde styr på task panes for hver inspector
        private Dictionary<Outlook.Inspector, CustomTaskPane> taskPanes =
            new Dictionary<Outlook.Inspector, CustomTaskPane>();

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Tilmeld NewInspector event - kaldes når et nyt vindue åbnes
            this.Application.Inspectors.NewInspector +=
                new Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);

            // Håndter eksisterende åbne vinduer
            foreach (Outlook.Inspector inspector in this.Application.Inspectors)
            {
                Inspectors_NewInspector(inspector);
            }
        }

        private void Inspectors_NewInspector(Outlook.Inspector inspector)
        {
            // Tjek om det er en appointment
            Outlook.AppointmentItem appointmentItem = inspector.CurrentItem as Outlook.AppointmentItem;

            if (appointmentItem != null)
            {
                // Opret custom task pane for denne inspector
                MeetingTaskPane taskPaneControl = new MeetingTaskPane();

                // Send appointment item til task pane så den kan indsætte skabelonen
                taskPaneControl.CurrentAppointment = appointmentItem;

                // Opret custom task pane
                CustomTaskPane taskPane = this.CustomTaskPanes.Add(
                    taskPaneControl,
                    "Møde Skabelon",
                    inspector
                );

                // Vis task pane automatisk!
                taskPane.Visible = true;
                taskPane.Width = 350;

                // Gem reference til task pane
                taskPanes[inspector] = taskPane;

                // Tilmeld inspector close event for at rydde op
                ((Outlook.InspectorEvents_10_Event)inspector).Close +=
                    new Outlook.InspectorEvents_10_CloseEventHandler(() => Inspector_Close(inspector));
            }
        }

        private void Inspector_Close(Outlook.Inspector inspector)
        {
            // Fjern task pane når inspector lukkes
            if (taskPanes.ContainsKey(inspector))
            {
                taskPanes.Remove(inspector);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Ryd op ved shutdown
            taskPanes.Clear();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
