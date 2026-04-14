export const LAUNCHES_SEED = [
  {
    id: "LN-2026-001",
    title: "Coronado Hotel delivery coordination",
    category: "Delivery",
    entryDate: "2026-04-09",
    client: "Coronado Hotel",
    location: "Jersey City, NJ",
    status: "Open",
    amount: 12480,
    summary:
      "Delivery coordination package for millwork arrival, unloading crew, and site staging before the installation team starts.",
    notes:
      "Validate loading dock access and keep the foreman aligned with the unloading window before arrival.",
    lineItems: [
      { id: "LI-01", description: "Unload crew", owner: "Warehouse Team", quantity: 2, unitPrice: 580 },
      { id: "LI-02", description: "Truck dispatch window", owner: "Logistics", quantity: 1, unitPrice: 940 },
      { id: "LI-03", description: "Temporary staging protection", owner: "Site Prep", quantity: 6, unitPrice: 180 }
    ]
  },
  {
    id: "LN-2026-002",
    title: "Madison trim installation phase",
    category: "Installation",
    entryDate: "2026-04-10",
    client: "Madison Residence",
    location: "Short Hills, NJ",
    status: "Ready",
    amount: 9860,
    summary:
      "Interior trim phase covering crown, casings, and punch list adjustments before paint touch-up.",
    notes:
      "Keep the second floor sequence first. Confirm stain samples before releasing the mantel section.",
    lineItems: [
      { id: "LI-01", description: "Trim carpenters", owner: "Finish Crew", quantity: 3, unitPrice: 760 },
      { id: "LI-02", description: "Punch corrections", owner: "Foreman", quantity: 5, unitPrice: 210 },
      { id: "LI-03", description: "Material restock", owner: "Procurement", quantity: 2, unitPrice: 340 }
    ]
  },
  {
    id: "LN-2026-003",
    title: "Brooklyn staircase field changes",
    category: "Field Change",
    entryDate: "2026-04-11",
    client: "Brooklyn Townhouse",
    location: "Brooklyn, NY",
    status: "In review",
    amount: 7340,
    summary:
      "Field adjustment entry for handrail return change, tread replacement, and revised baluster spacing.",
    notes:
      "Owner approved the sketch verbally. Formal confirmation still pending for the additional walnut finish.",
    lineItems: [
      { id: "LI-01", description: "Stair rework labor", owner: "Stair Crew", quantity: 2, unitPrice: 880 },
      { id: "LI-02", description: "Walnut treads", owner: "Shop", quantity: 4, unitPrice: 420 },
      { id: "LI-03", description: "Site verification visit", owner: "Project Manager", quantity: 1, unitPrice: 260 }
    ]
  },
  {
    id: "LN-2026-004",
    title: "Hudson warehouse receiving batch",
    category: "Receiving",
    entryDate: "2026-04-12",
    client: "Hudson Warehouse",
    location: "Newark, NJ",
    status: "Closed",
    amount: 5680,
    summary:
      "Receiving batch for prefinished cabinets, protective inventory check, and tagging before delivery allocation.",
    notes:
      "All crates matched the manifest. One drawer front was flagged for replacement and moved to hold.",
    lineItems: [
      { id: "LI-01", description: "Receiving labor", owner: "Warehouse Team", quantity: 4, unitPrice: 320 },
      { id: "LI-02", description: "Inspection and tagging", owner: "QC", quantity: 3, unitPrice: 190 },
      { id: "LI-03", description: "Hold shelf reservation", owner: "Inventory", quantity: 2, unitPrice: 120 }
    ]
  },
  {
    id: "LN-2026-005",
    title: "Palm Beach custom bar punch",
    category: "Punch List",
    entryDate: "2026-04-14",
    client: "Palm Beach Condo",
    location: "Palm Beach, FL",
    status: "Open",
    amount: 8420,
    summary:
      "Punch and completion entry for custom bar panel alignment, LED fix, and touch-up list before turnover.",
    notes:
      "Coordinate with electrician before closing the service panel. Client wants final photos the same day.",
    lineItems: [
      { id: "LI-01", description: "Punch crew", owner: "Finish Crew", quantity: 2, unitPrice: 640 },
      { id: "LI-02", description: "LED coordination", owner: "Electrical Partner", quantity: 1, unitPrice: 520 },
      { id: "LI-03", description: "Turnover detailing", owner: "Project Lead", quantity: 4, unitPrice: 230 }
    ]
  }
];
