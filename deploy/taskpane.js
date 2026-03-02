/*
 * Taskpane.js - UI funktionalitet
 */

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

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log("Taskpane loaded successfully");
  }
});

function insertTemplate() {
  Office.context.mailbox.item.body.setAsync(
    MEETING_TEMPLATE,
    { coercionType: Office.CoercionType.Html },
    function(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        alert("Skabelon indsat!");
      } else {
        alert("Fejl ved indsættelse: " + result.error.message);
      }
    }
  );
}
