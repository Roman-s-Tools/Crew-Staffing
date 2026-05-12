const fields = [
  "incidentName",
  "dateFrom",
  "dateTo",
  "timeFrom",
  "timeTo",

  "incidentCommander",
  "deputyIncidentCommander",
  "safetyOfficer",
  "publicInformationOfficer",
  "liaisonOfficer",

  "agency1",
  "agencyRep1",
  "agency2",
  "agencyRep2",
  "agency3",
  "agencyRep3",
  "agency4",
  "agencyRep4",
  "agency5",
  "agencyRep5",

  "planningChief",
  "planningDeputy",
  "resourcesUnit",
  "situationUnit",
  "documentationUnit",
  "demobilizationUnit",
  "technicalSpecialists",

  "logisticsChief",
  "logisticsDeputy",
  "supportBranchDirector",
  "supplyUnit",
  "facilitiesUnit",
  "groundSupportUnit",
  "serviceBranchDirector",
  "communicationsUnit",
  "medicalUnit",
  "foodUnit",

  "operationsChief",
  "operationsDeputy",
  "stagingArea",
  "branchDirector1",
  "branchDeputy1",
  "divisionGroup1",
  "divisionGroupSupervisor1",
  "divisionGroup2",
  "divisionGroupSupervisor2",
  "divisionGroup3",
  "divisionGroupSupervisor3",
  "branchDirector2",
  "branchDeputy2",
  "divisionGroup4",
  "divisionGroupSupervisor4",
  "divisionGroup5",
  "divisionGroupSupervisor5",
  "airOperationsBranchDirector",

  "financeChief",
  "financeDeputy",
  "timeUnit",
  "procurementUnit",
  "compClaimsUnit",
  "costUnit",

  "preparedByName",
  "preparedByTitle",
  "preparedBySignature",
  "preparedDateTime"
];

function getValue(id) {
  const element = document.getElementById(id);
  return element ? element.value.trim() : "";
}

function saveDraft() {
  const data = {};

  fields.forEach((field) => {
    data[field] = getValue(field);
  });

  localStorage.setItem("ics203Draft", JSON.stringify(data));
  alert("Draft saved.");
}

function loadDraft() {
  const saved = localStorage.getItem("ics203Draft");

  if (!saved) return;

  try {
    const data = JSON.parse(saved);

    fields.forEach((field) => {
      const element = document.getElementById(field);

      if (element && data[field]) {
        element.value = data[field];
      }
    });
  } catch (error) {
    console.error("Draft load error:", error);
  }
}

function clearForm() {
  if (!confirm("Clear all fields?")) return;

  fields.forEach((field) => {
    const element = document.getElementById(field);

    if (element) {
      element.value = "";
    }
  });

  localStorage.removeItem("ics203Draft");
}

function formatDocxError(error) {
  if (error.properties && error.properties.errors) {
    return error.properties.errors
      .map((err) => {
        return (
          err.properties?.explanation ||
          err.properties?.id ||
          err.message ||
          "Unknown template error"
        );
      })
      .join("\n\n");
  }

  return error.message || "Unknown export error";
}

async function exportICS203() {
  try {
    if (typeof PizZip === "undefined") {
      throw new Error("PizZip failed to load.");
    }

    if (typeof window.docxtemplater === "undefined") {
      throw new Error("Docxtemplater failed to load.");
    }

    if (typeof saveAs === "undefined") {
      throw new Error("FileSaver failed to load.");
    }

    const response = await fetch("/ics203.docx");

    if (!response.ok) {
      throw new Error(
        `Unable to load ics203.docx (HTTP ${response.status})`
      );
    }

    const arrayBuffer = await response.arrayBuffer();
    const zip = new PizZip(arrayBuffer);

    const doc = new window.docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: {
        start: "[[",
        end: "]]"
      }
    });

    const data = {};

    fields.forEach((field) => {
      data[field] = getValue(field);
    });

    doc.render(data);

    const blob = doc.getZip().generate({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });

    saveAs(blob, "ICS-203.docx");
  } catch (error) {
    console.error("ICS 203 Export Error:", error);

    alert("Export failed:\n\n" + formatDocxError(error));
  }
}

document.addEventListener("DOMContentLoaded", () => {
  loadDraft();

  const exportBtn = document.getElementById("exportBtn");
  const saveBtn = document.getElementById("saveBtn");
  const clearBtn = document.getElementById("clearBtn");

  if (exportBtn) exportBtn.addEventListener("click", exportICS203);
  if (saveBtn) saveBtn.addEventListener("click", saveDraft);
  if (clearBtn) clearBtn.addEventListener("click", clearForm);
});
