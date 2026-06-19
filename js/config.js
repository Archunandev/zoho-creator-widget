/**
 * FORM CONFIGURATION
 * ==================
 * Edit this array to match your Zoho Creator app's forms and fields.
 *
 * Each form object:
 *   id        — unique key used internally (any string)
 *   name      — display name shown in the dropdown
 *   linkName  — the form's link name in Zoho Creator (exact, case-sensitive)
 *   description — optional short description shown in step 1
 *   fields[]  — list of fields to import into
 *
 * Each field object:
 *   label     — display label shown in the UI
 *   linkName  — the field's link name in Zoho Creator (exact, case-sensitive)
 *   required  — true = validation error if empty
 *   type      — "text" | "email" | "number" | "date" | "url" | "phone"
 */
const FORM_CONFIG = [
  {
    id: "Pincode",
    name: "Pincode",
    linkName: "Pincode_Master",
    description: "Pincode Master records",
    fields: [
      { label: "Pincode",   linkName: "Pincode",   required: true,  type: "text"   },
      { label: "Country",    linkName: "Country",    required: true,  type: "text"   },
      { label: "District",        linkName: "District",        required: true,  type: "text"  },
    ]
  }
];
