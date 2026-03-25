export default function SuperBuilderSection() {
  return (
    <section id="recommendation" className="border-b border-[#1f1f1f]">
      {/* Section Header */}
      <div className="max-w-[1280px] mx-auto px-6 py-16 border-b border-[#1f1f1f]">
        <div className="flex items-center gap-4">
          <span className="font-mono text-[11px] tracking-[0.25em] text-[#3B82F6] uppercase">
            05 / Recommendation
          </span>
          <div className="h-[1px] flex-1 bg-[#1f1f1f]" />
        </div>
        <h2 className="font-display text-[clamp(48px,6vw,96px)] leading-[0.95] text-[#f0f0f0] uppercase mt-6">
          The Super Builder
          <br />
          <span className="text-[#3B82F6]">Stack</span>
        </h2>
        <p className="font-sans text-[15px] text-[#888] mt-4 max-w-2xl">
          For product builders at Motive Technologies: the optimal two-tool combo for rapid prototyping,
          validation, and production deployment.
        </p>
      </div>

      <div className="max-w-[1280px] mx-auto px-6 py-16">
        {/* Main Recommendation */}
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 mb-12">
          {/* Claude Code */}
          <div className="group relative border border-[#1f1f1f] hover:border-[#3B82F6] bg-[#0f0f0f] p-8 transition-all duration-300">
            <div className="absolute top-0 left-0 w-full h-[3px] bg-gradient-to-r from-[#D97706] to-[#F59E0B] opacity-0 group-hover:opacity-100 transition-opacity duration-300" />

            <div className="flex items-start justify-between mb-6">
              <div className="w-14 h-14 flex items-center justify-center border border-[#2e2e2e] font-display text-[28px] leading-none text-[#D97706]">
                A
              </div>
              <span className="font-mono text-[10px] tracking-[0.15em] text-[#D97706] border border-[#D97706]/30 bg-[#D97706]/10 px-3 py-1 uppercase">
                Terminal
              </span>
            </div>

            <h3 className="font-display text-[42px] leading-none text-[#f0f0f0] uppercase mb-4">
              Claude Code
            </h3>

            <p className="font-sans text-[15px] leading-[1.7] text-[#888] mb-6">
              Terminal-based agentic coding for complex production codebases. Understands entire projects,
              makes multi-file changes, runs tests, and handles architecture decisions.
            </p>

            <div className="space-y-3 mb-6">
              <div className="flex items-start gap-3">
                <div className="w-5 h-5 rounded-full bg-[#D97706]/20 flex items-center justify-center flex-shrink-0 mt-0.5">
                  <span className="text-[#D97706] text-[11px]">✓</span>
                </div>
                <p className="font-sans text-[14px] text-[#aaa]">
                  <strong className="text-[#f0f0f0]">Production-grade refactoring:</strong> Handle large-scale
                  changes across entire codebases with confidence
                </p>
              </div>
              <div className="flex items-start gap-3">
                <div className="w-5 h-5 rounded-full bg-[#D97706]/20 flex items-center justify-center flex-shrink-0 mt-0.5">
                  <span className="text-[#D97706] text-[11px]">✓</span>
                </div>
                <p className="font-sans text-[14px] text-[#aaa]">
                  <strong className="text-[#f0f0f0]">1M token context window:</strong> Full project understanding
                  for accurate, context-aware changes
                </p>
              </div>
              <div className="flex items-start gap-3">
                <div className="w-5 h-5 rounded-full bg-[#D97706]/20 flex items-center justify-center flex-shrink-0 mt-0.5">
                  <span className="text-[#D97706] text-[11px]">✓</span>
                </div>
                <p className="font-sans text-[14px] text-[#aaa]">
                  <strong className="text-[#f0f0f0]">Extended thinking mode:</strong> Handles complex
                  architecture and debugging tasks autonomously
                </p>
              </div>
            </div>

            <div className="border-t border-[#1f1f1f] pt-5">
              <div className="flex items-center justify-between">
                <div>
                  <div className="font-mono text-[11px] tracking-[0.1em] text-[#3B82F6] uppercase">
                    From $20/mo
                  </div>
                  <div className="font-sans text-[13px] text-[#666] mt-1">
                    Start with Pro plan and expand on requirement
                  </div>
                </div>
                <div className="text-right">
                  <div className="font-display text-[24px] leading-none text-[#D97706]">
                    01
                  </div>
                  <div className="font-mono text-[10px] text-[#444] uppercase tracking-wider">
                    Production
                  </div>
                </div>
              </div>
            </div>
          </div>

          {/* Lovable */}
          <div className="group relative border border-[#1f1f1f] hover:border-[#3B82F6] bg-[#0f0f0f] p-8 transition-all duration-300">
            <div className="absolute top-0 left-0 w-full h-[3px] bg-gradient-to-r from-[#EC4899] to-[#F472B6] opacity-0 group-hover:opacity-100 transition-opacity duration-300" />

            <div className="flex items-start justify-between mb-6">
              <div className="w-14 h-14 flex items-center justify-center border border-[#2e2e2e] font-display text-[28px] leading-none text-[#EC4899]">
                L
              </div>
              <span className="font-mono text-[10px] tracking-[0.15em] text-[#22c55e] border border-[#22c55e]/30 bg-[#22c55e]/10 px-3 py-1 uppercase">
                Builder
              </span>
            </div>

            <h3 className="font-display text-[42px] leading-none text-[#f0f0f0] uppercase mb-4">
              Lovable
            </h3>

            <p className="font-sans text-[15px] leading-[1.7] text-[#888] mb-6">
              Full-stack app builder that turns natural language into React + Supabase applications. Perfect for
              rapid prototyping, MVP validation, and product exploration.
            </p>

            <div className="space-y-3 mb-6">
              <div className="flex items-start gap-3">
                <div className="w-5 h-5 rounded-full bg-[#EC4899]/20 flex items-center justify-center flex-shrink-0 mt-0.5">
                  <span className="text-[#EC4899] text-[11px]">✓</span>
                </div>
                <p className="font-sans text-[14px] text-[#aaa]">
                  <strong className="text-[#f0f0f0]">Zero-to-prototype in hours:</strong> Validate product
                  concepts before committing engineering resources
                </p>
              </div>
              <div className="flex items-start gap-3">
                <div className="w-5 h-5 rounded-full bg-[#EC4899]/20 flex items-center justify-center flex-shrink-0 mt-0.5">
                  <span className="text-[#EC4899] text-[11px]">✓</span>
                </div>
                <p className="font-sans text-[14px] text-[#aaa]">
                  <strong className="text-[#f0f0f0]">Figma-like visual editor:</strong> Non-technical PMs can
                  tweak UI without touching code
                </p>
              </div>
              <div className="flex items-start gap-3">
                <div className="w-5 h-5 rounded-full bg-[#EC4899]/20 flex items-center justify-center flex-shrink-0 mt-0.5">
                  <span className="text-[#EC4899] text-[11px]">✓</span>
                </div>
                <p className="font-sans text-[14px] text-[#aaa]">
                  <strong className="text-[#f0f0f0]">GitHub export:</strong> Own your code — graduate validated
                  prototypes to production
                </p>
              </div>
            </div>

            <div className="border-t border-[#1f1f1f] pt-5">
              <div className="flex items-center justify-between">
                <div>
                  <div className="font-mono text-[11px] tracking-[0.1em] text-[#3B82F6] uppercase">
                    From $25/mo
                  </div>
                  <div className="font-sans text-[13px] text-[#666] mt-1">
                    Start with Business Plan and then scale with enterprise post analyzing usage
                  </div>
                </div>
                <div className="text-right">
                  <div className="font-display text-[24px] leading-none text-[#EC4899]">
                    02
                  </div>
                  <div className="font-mono text-[10px] text-[#444] uppercase tracking-wider">
                    Prototype
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Why This Combo */}
        <div className="border border-[#3B82F6]/30 bg-gradient-to-br from-[#0f0f0f] to-[#161616] p-10">
          <div className="flex items-start gap-6">
            <div className="w-16 h-16 rounded-full bg-[#3B82F6]/20 flex items-center justify-center flex-shrink-0">
              <span className="font-display text-[32px] leading-none text-[#3B82F6]">+</span>
            </div>
            <div className="flex-1">
              <h3 className="font-display text-[32px] leading-none text-[#f0f0f0] uppercase mb-4">
                Why This Combo Works
              </h3>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div>
                  <h4 className="font-sans text-[15px] font-semibold text-[#3B82F6] mb-2 uppercase tracking-wide">
                    For Product Validation
                  </h4>
                  <p className="font-sans text-[14px] leading-[1.7] text-[#888]">
                    Start in <strong className="text-[#f0f0f0]">Lovable</strong> — build clickable prototypes in
                    hours, test with users, iterate on UX. Export clean code when validated. No engineering
                    bottleneck, no throwaway work.
                  </p>
                </div>
                <div>
                  <h4 className="font-sans text-[15px] font-semibold text-[#3B82F6] mb-2 uppercase tracking-wide">
                    For Production Polish
                  </h4>
                  <p className="font-sans text-[14px] leading-[1.7] text-[#888]">
                    Graduate to <strong className="text-[#f0f0f0]">Claude Code</strong> — refactor for scale,
                    add complex features, integrate with existing systems. Terminal-based workflow integrates with
                    any development environment.
                  </p>
                </div>
              </div>
              <div className="mt-6 pt-6 border-t border-[#1f1f1f]">
                <p className="font-mono text-[13px] text-[#3B82F6] uppercase tracking-wide mb-2">
                  Recommended Workflow
                </p>
                <div className="flex items-center gap-4 flex-wrap">
                  <div className="flex items-center gap-2">
                    <span className="font-display text-[18px] text-[#EC4899]">Lovable</span>
                    <span className="text-[#444]">→</span>
                    <span className="font-sans text-[13px] text-[#888]">Prototype + Validate</span>
                  </div>
                  <span className="text-[#444] hidden md:inline">→</span>
                  <div className="flex items-center gap-2">
                    <span className="font-display text-[18px] text-[#D97706]">Claude Code</span>
                    <span className="text-[#444]">→</span>
                    <span className="font-sans text-[13px] text-[#888]">Refactor + Ship</span>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </section>
  );
}
