export interface UseCase {
  role: string;
  tools: string[];
  description: string;
}

export const useCases: UseCase[] = [
  {
    role: 'Non-Technical PM',
    tools: ['Lovable', 'Bolt'],
    description: 'Build internal dashboards, clickable prototypes, and validate product ideas without writing a line of code.',
  },
  {
    role: 'Frontend Engineer',
    tools: ['v0', 'Cursor'],
    description: 'Generate React component scaffolding with v0, then refine and extend with Cursor for production-quality code.',
  },
  {
    role: 'Full-Stack Engineer',
    tools: ['Cursor', 'Windsurf', 'Claude Code'],
    description: 'Multi-file refactors, architecture decisions, and complex feature development across large codebases.',
  },
  {
    role: 'Rapid Prototyper',
    tools: ['Bolt', 'Lovable', 'v0'],
    description: 'Zero-to-working-demo in under an hour. Validate UX assumptions before any engineering investment.',
  },
  {
    role: 'Enterprise / Production',
    tools: ['Cursor', 'GitHub Copilot', 'Claude Code'],
    description: 'IP indemnity, SOC 2 compliance, audit logs, and enterprise SSO. Safe for production-critical workflows.',
  },
  {
    role: 'Learning / Experimentation',
    tools: ['Replit', 'Lovable'],
    description: 'Explore new frameworks, practice new languages, and build proof-of-concepts without any local environment setup.',
  },
];
