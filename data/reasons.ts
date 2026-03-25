export interface Reason {
  number: string;
  title: string;
  description: string;
}

export const reasons: Reason[] = [
  {
    number: '01',
    title: 'Speed to Validation',
    description: 'Compress concept-to-clickable-prototype from weeks to hours. Validate product ideas before committing development resources.',
  },
  {
    number: '02',
    title: 'Product Team Autonomy',
    description: 'PMs and designers build functional prototypes, test UX concepts, and create internal tools without engineering dependencies.',
  },
  {
    number: '03',
    title: 'Quality at Scale',
    description: 'AI tools enforce best practices, generate production-ready code, and maintain consistency across rapid iterations.',
  },
  {
    number: '04',
    title: 'Competitive Edge',
    description: 'YC W25 data: 21% of startups run 95%+ AI-generated codebases. Product builders at Motive need fluency in vibe coding.',
  },
  {
    number: '05',
    title: 'Reduce Validation Costs',
    description: 'Test product hypotheses in Lovable or Bolt, graduate validated concepts to production, and avoid building the wrong thing.',
  },
];
