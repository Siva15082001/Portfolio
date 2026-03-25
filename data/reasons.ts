export interface Reason {
  number: string;
  title: string;
  description: string;
}

export const reasons: Reason[] = [
  {
    number: '01',
    title: 'Speed to Signal',
    description: 'Compress prototype-to-demo from weeks to hours. Ship faster, learn faster, iterate before committing engineering resources.',
  },
  {
    number: '02',
    title: 'PM & Design Autonomy',
    description: 'Non-engineers build internal tools and product explorations without consuming engineering bandwidth.',
  },
  {
    number: '03',
    title: 'Quality at Scale',
    description: 'AI IDEs enforce consistent architecture patterns, generate tests automatically, and flag regressions before they ship.',
  },
  {
    number: '04',
    title: 'Talent Density',
    description: 'YC W25 data: 21% of startups run 95%+ AI-generated codebases. Motive engineers need fluency in this paradigm.',
  },
  {
    number: '05',
    title: 'Reduce Prototype Debt',
    description: 'Validate ideas in Lovable or Bolt, graduate clean code to Cursor, and avoid the costly full-rewrite cycle.',
  },
];
