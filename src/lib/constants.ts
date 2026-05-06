// ── Business Constants ─────────────────────────────────────────────────────

export const RATES = {
  FSM_I: { regular: 31.05, ot: 15.53 },
  FSM_II: { regular: 35.50, ot: 17.75 },
} as const;

export const MARKUP = {
  FT: 0.2993,
  PT: 0.2770,
} as const;

export const OT_COMMENT_PATTERNS = ['over time', 'ca daily ot'];

export const FSM_PROGRAM_PATTERNS = ['fsm', 'samsung field sales manager'];

export interface MgmtMember {
  name: string;
  associateId: string;
  title: string;
  associateState: string;
  hourlyRate: number;
  allocationPct: number; // 0-1
}

export const MGMT_TABLE: MgmtMember[] = [
  { name: 'Mike King',                  associateId: 'MK1002A',   title: 'Account Director',               associateState: 'TX', hourlyRate: 92.32, allocationPct: 0.3682 },
  { name: 'Corey Purdin',               associateId: 'CP1010A',   title: 'National Manager',                associateState: 'TX', hourlyRate: 91.97, allocationPct: 0.3682 },
  { name: 'Elizabeth Hobson',           associateId: 'EW1002I',   title: 'National Manager, SEC',           associateState: 'OR', hourlyRate: 85.46, allocationPct: 0.0000 },
  { name: 'Shawn Scialo',               associateId: 'SS1183I',   title: 'National Manager, FSM',           associateState: 'AZ', hourlyRate: 85.46, allocationPct: 1.0000 },
  { name: 'Mike Coronado',              associateId: 'EE016130',  title: 'Operations Manager',              associateState: 'TX', hourlyRate: 71.24, allocationPct: 0.4000 },
  { name: 'Kristen Eberline',           associateId: 'KE1021I',   title: 'Project Manager',                 associateState: 'KS', hourlyRate: 63.39, allocationPct: 0.4000 },
  { name: 'Lisa Ghattas',               associateId: 'LJ1011I',   title: 'Operations Coordinator 1',       associateState: 'MO', hourlyRate: 38.75, allocationPct: 0.5700 },
  { name: 'Mason Vanmeter',             associateId: 'MV1046I',   title: 'Operations Coordinator 2',       associateState: 'AZ', hourlyRate: 38.75, allocationPct: 0.4700 },
  { name: 'George Macias',              associateId: 'JM1347I',   title: 'Data Analyst',                    associateState: 'TX', hourlyRate: 63.39, allocationPct: 0.5000 },
  { name: 'William Gobar',              associateId: 'WG1011I',   title: 'Insights Analyst, SEC',           associateState: 'MD', hourlyRate: 47.54, allocationPct: 0.0000 },
  { name: 'Amanda Bradshaw',            associateId: 'EE004278',  title: 'Onboarding Specialist, FSM',      associateState: 'TN', hourlyRate: 47.54, allocationPct: 1.0000 },
  { name: 'Patrick Mendez',             associateId: 'PM1024I',   title: 'Onboarding Specialist, SEC',      associateState: 'CA', hourlyRate: 47.54, allocationPct: 0.0000 },
  { name: 'Mary Beth French',           associateId: 'MF1010I',   title: 'Training Manager',                associateState: 'TN', hourlyRate: 54.92, allocationPct: 0.3700 },
  { name: 'Hunter West',                associateId: 'HW1020I',   title: 'Field Operations Manager I',      associateState: 'AZ', hourlyRate: 53.88, allocationPct: 0.4300 },
  { name: 'Yamil Saade',                associateId: 'YS1013I',   title: 'Field Operations Manager I',      associateState: 'TX', hourlyRate: 53.88, allocationPct: 0.4300 },
  { name: 'David Exum',                 associateId: 'DE1016I',   title: 'Field Operations Manager I',      associateState: 'CA', hourlyRate: 54.93, allocationPct: 1.0000 },
  { name: 'Eric Lopez',                 associateId: 'EL1011I',   title: 'Field Operations Manager I',      associateState: 'KY', hourlyRate: 53.88, allocationPct: 0.4300 },
  { name: 'Matthew Wickham',            associateId: 'MW1112I',   title: 'Field Operations Manager I',      associateState: 'VA', hourlyRate: 53.88, allocationPct: 0.4300 },
  { name: 'Miguel Aguilar',             associateId: 'MA1183C',   title: 'Field Operations Manager I',      associateState: 'VA', hourlyRate: 53.88, allocationPct: 0.4300 },
  { name: 'Sara Wali',                  associateId: 'SW1147C',   title: 'Field Operations Manager I',      associateState: 'CA', hourlyRate: 53.88, allocationPct: 0.4300 },
  { name: 'Steffanie Molina-Frybarger', associateId: 'EE003697',  title: 'Field Operations Manager I',      associateState: 'TX', hourlyRate: 53.88, allocationPct: 0.4300 },
  { name: 'Donald Scarfo',              associateId: 'EE010525',  title: 'Field Operations Manager I',      associateState: 'SC', hourlyRate: 53.88, allocationPct: 0.4300 },
  { name: 'David Akom',                 associateId: 'EE003625',  title: 'Field Operations Manager I',      associateState: 'GA', hourlyRate: 52.31, allocationPct: 0.4300 },
  { name: 'Rebecca Tejeda',             associateId: 'EE010554',  title: 'Field Operations Manager I',      associateState: 'IL', hourlyRate: 52.31, allocationPct: 0.4300 },
  { name: 'Kenneth Hitt',               associateId: 'EE010650',  title: 'Field Operations Manager I',      associateState: 'AZ', hourlyRate: 53.88, allocationPct: 0.4300 },
  { name: 'Belinda Chi',                associateId: 'EE010255',  title: 'Field Operations Manager I',      associateState: 'WA', hourlyRate: 52.31, allocationPct: 0.4300 },
  { name: 'Brian Conner',               associateId: 'EE002390',  title: 'Inventory Specialist',            associateState: 'TX', hourlyRate: 37.63, allocationPct: 0.5000 },
  { name: 'Billy MacDonald',            associateId: 'WM1054A',   title: 'Inventory/IT Specialist',         associateState: 'TX', hourlyRate: 37.48, allocationPct: 1.0000 },
];
