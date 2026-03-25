export type ToolCategory = 'IDE' | 'Builder' | 'Prototyping' | 'Agent' | 'UI Generator' | 'Terminal';

export interface ToolFeature {
  text: string;
}

export interface Tool {
  id: string;
  name: string;
  tagline: string;
  category: ToolCategory;
  features: string[];
  freeTier: string;
  proPriceMonthly: string;
  businessPrice: string;
  bestFor: string;
  aiModels: string;
  fullStack: 'yes' | 'no' | 'partial';
  githubSync: 'yes' | 'no' | 'partial';
  deployment: string;
  letterMark: string;
  accentColor: string;
  verdict: string;
}

export const tools: Tool[] = [
  {
    id: 'cursor',
    name: 'Cursor',
    tagline: 'AI-native IDE with deep codebase context and multi-file agents.',
    category: 'IDE',
    features: [
      'Codebase-wide awareness via RAG-based indexing',
      'Agent mode for autonomous multi-file edits',
      'Tab autocomplete with next-edit prediction',
      'Multi-model: Claude, GPT-4o, Gemini — credit-based',
      'Privacy mode with zero data retention options',
    ],
    freeTier: 'Hobby (limited agent requests)',
    proPriceMonthly: '$20/mo',
    businessPrice: 'Teams $40/user/mo',
    bestFor: 'Full-stack engineers in daily development',
    aiModels: 'Claude Sonnet 4.6, GPT-4o, Gemini',
    fullStack: 'yes',
    githubSync: 'yes',
    deployment: 'Via integrated terminal',
    letterMark: 'C',
    accentColor: '#7C3AED',
    verdict: 'Best daily driver IDE for engineers. Credit-based model warrants usage tracking.',
  },
  {
    id: 'lovable',
    name: 'Lovable',
    tagline: 'Full-stack app builder. Describe it, it builds it, ships it.',
    category: 'Builder',
    features: [
      'Natural language to React + Supabase full-stack apps',
      'Figma-like visual editor for non-technical stakeholders',
      'GitHub export at any time — own your code fully',
      'Security scan on publish + Supabase native integration',
      'Custom domains — 10,000+ connected since launch',
    ],
    freeTier: '5 daily credits, 20 collaborators',
    proPriceMonthly: '$25/mo',
    businessPrice: 'Business $50/mo',
    bestFor: 'PMs and non-engineers building MVPs',
    aiModels: 'Anthropic Claude (internal)',
    fullStack: 'yes',
    githubSync: 'yes',
    deployment: 'Lovable Cloud + Vercel/Netlify export',
    letterMark: 'L',
    accentColor: '#EC4899',
    verdict: 'Best for MVP validation. Credit burn unpredictable on complex projects.',
  },
  {
    id: 'bolt',
    name: 'Bolt',
    tagline: 'Browser-based agentic builder powered by StackBlitz WebContainers.',
    category: 'Prototyping',
    features: [
      'Full-stack apps run entirely in-browser via WebContainers',
      'Agentic iteration — auto-fixes build errors without prompting',
      'Figma import: drop designs into chat for code generation',
      'Claude Opus 4.6 and lightweight model toggle for speed/cost',
      'Open source engine — self-hostable core via bolt.diy',
    ],
    freeTier: '1M tokens/mo, 300K daily',
    proPriceMonthly: '$20/mo',
    businessPrice: 'Teams $30/user/mo',
    bestFor: 'Rapid prototyping without local setup',
    aiModels: 'Claude Opus 4.6, GPT-4o',
    fullStack: 'partial',
    githubSync: 'yes',
    deployment: 'Netlify (native), export to any host',
    letterMark: 'B',
    accentColor: '#0EA5E9',
    verdict: 'Fastest zero-to-prototype. Token budget depletes rapidly on heavy projects.',
  },
  {
    id: 'v0',
    name: 'v0',
    tagline: 'React + shadcn/ui component generator by the Vercel team.',
    category: 'UI Generator',
    features: [
      'Generates React/Next.js + Tailwind + shadcn/ui components',
      'One-click deploy to Vercel infrastructure',
      'Iterative chat-based refinement with visual previews',
      'Code execution in JS and Python, diagram generation',
      'Full-stack context from Vercel — knows deployment best practices',
    ],
    freeTier: '$5/mo in credits (resets monthly)',
    proPriceMonthly: '$20/mo',
    businessPrice: '$100/user/mo',
    bestFor: 'Frontend engineers building UI components fast',
    aiModels: 'Vercel internal models (GPT-4 class)',
    fullStack: 'no',
    githubSync: 'yes',
    deployment: 'Vercel (one-click native)',
    letterMark: 'V',
    accentColor: '#000000',
    verdict: 'Unmatched for React/shadcn UI. Limited backend. Vercel dependency.',
  },
  {
    id: 'replit',
    name: 'Replit',
    tagline: 'Cloud IDE with AI Agent that codes, deploys, and hosts in one browser tab.',
    category: 'Agent',
    features: [
      'Agent 3 with effort-based pricing — pays per task complexity',
      '50+ languages, zero local setup required',
      'Built-in hosting and database — ship without leaving the IDE',
      'Turbo Mode on Pro: 2x speed with frontier model access',
      'Real-time multiplayer collaboration on any project',
    ],
    freeTier: 'Starter (free daily credits, 1 app)',
    proPriceMonthly: 'Core $20/mo',
    businessPrice: 'Pro $100/mo (15 builders)',
    bestFor: 'Learning, experimentation, rapid cloud prototyping',
    aiModels: 'Replit Agent models (Claude + custom)',
    fullStack: 'yes',
    githubSync: 'yes',
    deployment: 'Replit hosting (native)',
    letterMark: 'R',
    accentColor: '#F26207',
    verdict: 'Best all-in-one cloud IDE. Cost unpredictable with heavy Agent use.',
  },
  {
    id: 'windsurf',
    name: 'Windsurf',
    tagline: 'Agentic IDE with Cascade — understands your entire codebase.',
    category: 'IDE',
    features: [
      'Cascade: multi-step agentic assistant with deep repo context',
      'Unlimited Tab autocomplete on every plan including free',
      'BYOK support — use your own Anthropic/OpenAI API keys',
      'SWE-1 native model family optimized for software engineering',
      'SOC 2 Type II, FedRAMP High, JetBrains plugin support',
    ],
    freeTier: '25 monthly prompt credits, unlimited Tab',
    proPriceMonthly: '$15/mo',
    businessPrice: 'Teams $30/user/mo',
    bestFor: 'Engineers wanting Cursor alternative with lower price',
    aiModels: 'SWE-1, GPT-5, Claude, Gemini (BYOK)',
    fullStack: 'yes',
    githubSync: 'yes',
    deployment: 'Via integrated terminal + Netlify preview',
    letterMark: 'W',
    accentColor: '#06B6D4',
    verdict: 'Strongest Cursor alternative. $3B OpenAI acquisition validates quality.',
  },
  {
    id: 'copilot',
    name: 'GitHub Copilot',
    tagline: 'Inline AI completion with deep IDE and GitHub ecosystem integration.',
    category: 'IDE',
    features: [
      'Inline code completion with 2,000+ editor integrations',
      'Copilot Chat with agent mode for multi-file edits',
      'Deep GitHub context: PRs, issues, codebase knowledge bases',
      'Multi-model: Claude Sonnet 4.6, GPT-4.1, o3 on Pro+',
      'IP indemnity and enterprise-grade SSO on Business/Enterprise',
    ],
    freeTier: '2,000 completions + 50 chat/mo',
    proPriceMonthly: '$10/mo',
    businessPrice: 'Business $19/user/mo',
    bestFor: 'Enterprises needing safe, compliant AI coding',
    aiModels: 'Claude Sonnet 4.6, GPT-4.1, o3',
    fullStack: 'partial',
    githubSync: 'yes',
    deployment: 'Via IDE (no native deploy)',
    letterMark: 'G',
    accentColor: '#22C55E',
    verdict: 'Most cost-efficient for compliant teams. Best when deeply on GitHub.',
  },
  {
    id: 'claude-code',
    name: 'Claude Code',
    tagline: 'Terminal-based agentic coding for complex, large-scale codebases.',
    category: 'Terminal',
    features: [
      'Runs in your terminal — integrates with any editor or pipeline',
      'Full codebase understanding via 1M token context window',
      'Agentic execution: reads files, runs tests, makes commits',
      'Extended thinking for architecture and debugging tasks',
      'CLAUDE.md memory files for persistent project context',
    ],
    freeTier: 'None (requires Pro plan)',
    proPriceMonthly: '$20/mo (Pro)',
    businessPrice: 'Max $100/mo, Teams Premium $150/user/mo',
    bestFor: 'Senior engineers on large, complex production codebases',
    aiModels: 'Claude Opus 4.6, Claude Sonnet 4.6',
    fullStack: 'yes',
    githubSync: 'yes',
    deployment: 'Terminal-driven (CI/CD native)',
    letterMark: 'A',
    accentColor: '#D97706',
    verdict: 'Most powerful for complex codebases. Steeper learning curve, high value.',
  },
];
