/*
 * Commands.js - Event-Based Activation for Meeting Template
 * FIXED: Bruger OnNewAppointmentOrganizer event
 */

// Møde skabelon i HTML format
const MEETING_TEMPLATE = `
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
<p>Tid til spørgsmål eller andre punkter.</p>
`;

// Initialize Office.js
Office.onReady(() => {
  console.log('Office.js ready - Meeting Template Add-in loaded');
});

/**
 * Event handler for OnNewAppointmentOrganizer
 * Kaldes automatisk når en ny mødeaftale oprettes
 *
 * VIGTIGT: event.completed() SKAL kaldes når færdig!
 */
function onNewAppointmentOrganizer(event) {
  console.log('OnNewAppointmentOrganizer event triggered');

  const item = Office.context.mailbox.item;

  // Check om body allerede har indhold (undgå duplicate insert)
  item.body.getAsync(Office.CoercionType.Text, function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {

      const bodyText = result.value.trim();

      // Kun indsæt hvis body er tom eller næsten tom
      if (bodyText.length < 10) {

        console.log('Body is empty - inserting template');

        // Indsæt skabelon i møde body
        item.body.setAsync(
          MEETING_TEMPLATE,
          { coercionType: Office.CoercionType.Html },
          function(setResult) {
            if (setResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log('Template inserted successfully');
            } else {
              console.error('Error inserting template:', setResult.error);
            }

            // KRITISK: Signaler at event handler er færdig
            event.completed();
          }
        );

      } else {
        console.log('Body already has content - skipping template insert');
        // KRITISK: event.completed() SKAL kaldes selv hvis vi springer insert over
        event.completed();
      }

    } else {
      console.error('Error reading body:', result.error);
      // KRITISK: Kald event.completed() selv ved fejl
      event.completed();
    }
  });
}

// Registrer event handler globalt så manifest kan kalde den
// VIGTIGT: Function navn skal matche præcis det der står i manifest
if (typeof Office !== 'undefined' && Office.actions) {
  Office.actions.associate("onNewAppointmentOrganizer", onNewAppointmentOrganizer);
  console.log('Event handler registered: onNewAppointmentOrganizer');
}
