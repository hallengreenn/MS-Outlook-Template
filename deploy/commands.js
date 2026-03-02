/*
 * Commands.js - Håndterer automatisk indsættelse af møde skabelon
 */

// Skabelon tekst i HTML format med Neue Haas Grotesk font
const MEETING_TEMPLATE = `
<div style="font-family: 'Neue Haas Grotesk', Arial, sans-serif;">
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
</div>
`;

// Denne funktion kaldes automatisk når brugeren opretter en ny mødeaftale
function onNewAppointmentCompose(event) {
  const item = Office.context.mailbox.item;

  // Sæt skabelon i møde body
  item.body.setAsync(
    MEETING_TEMPLATE,
    { coercionType: Office.CoercionType.Html },
    function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Skabelon indsat succesfuldt");
      } else {
        console.error("Fejl ved indsættelse af skabelon:", result.error);
      }
      // Signaler at event handler er færdig
      event.completed();
    }
  );
}

// Registrer funktioner globalt så manifest kan kalde dem
if (typeof Office !== 'undefined') {
  Office.actions.associate("onNewAppointmentCompose", onNewAppointmentCompose);
}
