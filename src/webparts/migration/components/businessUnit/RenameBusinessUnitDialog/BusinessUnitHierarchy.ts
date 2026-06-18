export interface IBusinessUnitHierarchy {
  [businessUnit: string]: {
    [department: string]: string[];
  };
}

export const BUSINESS_UNIT_HIERARCHY: IBusinessUnitHierarchy = {
  'Leadership & Organization Development': {
    'Organization Development': [],
    iAcademy: []
  },
  'Chief Technology Office': {},
  'Commercial Technology Group': {
    'DAAI and Tech Solutions': [
      'Omnichannel & Medical Analytics',
      'Field, Customer & Enterprise Services',
      'Data Strategy, Patient & Clinical',
      'CDP, Tagging & Taxonomy',
      'Technical Solutions'
    ],
    Omnichannel: [
      'Orchestration',
      'Campaign and Media',
      'Omnichannel Activation'
    ],
    'Enterprise Content Operations': [
      'Commercial Delivery Services',
      'Web and Portal Solutions',
      'CXD Content and Experience Design',
      'Tectonic',
      'Patient Experience'
    ],
    'Commercial Excellence': [
      'Platforms'
    ]
  },
  'Activation Framework': {},
  'Amgen C 360': {},
  'Enterprise Medical': {
    'Global Safety': [
      'Safety',
      'Pharmacovigilance'
    ],
    'Medical Affairs': [
      'Medical Communication',
      'Promotional and Medical Review',
      'Material Review, Operations and Compliance'
    ],
    'Regulatory Solutions': [],
    'Packaging Artwork and Labelling': [],
    'PRMA & HEOR': [],
    'Enterprise Medical Platforms and Technology': []
  },
  'Device Technology - QARA': {},
  'Enterprise Clinical': {},
  'Strategy and Development': {},
  'Finance and Legal': {
    'Commercial & Facilities': [],
    'Risk Audit and Compliance': [
      'Process Management',
      'Third Party Risk Management',
      'Internal Audit'
    ]
  },
  'Global Operations': {
    'Knowledge Management': [],
    'Process and Automation': [],
    'Metrics & Governance': [],
    'Central Resource Group (CRG)': [],
    'Enterprise Applications': [],
    'Enterprise IT Infrastructure': []
  },
  'People Practices & Systems': {},
  'Sales Lifesciences': {
    'Sales China': [],
    'Growth and Emerging Accounts': [],
    'Key Accounts': []
  },
  'Global Delivery Center': {},
  'Indegene Products': {},
  Marketing: {},
  'DT Consulting': {},
  CultHealth: {}
};
