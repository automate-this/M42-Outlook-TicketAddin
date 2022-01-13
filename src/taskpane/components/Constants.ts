export const TicketTypes = {
    Ticket: "SPSActivityTypeTicket",
    Incident: "SPSActivityTypeIncident",
    ServiceRequest: "SPSActivityTypeServiceRequest"
} as const;

export const CustomPropertyNames = {
    M42TicketID: "M42TicketID",
    M42TicketNumber: "M42TicketNumber",
    M42TicketType: "M42TicketType"
} as const;

export const DefaultCategoryGUID = "42b49002-fed3-4c9b-9532-cf351df038cf" as const; //Service Desk > Störungen
export const DefaultInitiatorGUID = "5efd7997-9ca1-eb11-c591-0050569bdda3" as const; //"Dummy"-User