/*
  Optional name mapping. Add staff entries using their actual Microsoft 365 sign-in
  email address and display name. If no email match is found, the user can type
  their name manually in the add-in.

  This is a convenience lookup, not identity verification.
*/
window.BRIDMET_SIGNATURE_CONFIG = {
  users: [
    // { email: "employee@bridmet.co.uk", name: "Employee Name" }
  ],
  telephone: "01308 426444",
  website: "https://www.bridmet.co.uk",
  websiteLabel: "WWW.BRIDMET.CO.UK",
  locationLine1: "CROSS BUSINESS PARK, BRIDPORT,",
  locationLine2: "DORSET, DT6 3UX",
  certificationLines: [
    "Bridmet Ltd has been certified by BSI",
    "to ISO9001:2015, ISO14001:2015 &",
    "ISO45001:2018 under certificate",
    "numbers FM 588509, EMS 683665 &",
    "OHS 691239"
  ]
};
