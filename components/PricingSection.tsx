import { tools } from '@/data/tools';

export default function PricingSection() {
  const row1 = tools.slice(0, 4);
  const row2 = tools.slice(4, 8);
  const row3 = tools.slice(8, 12);

  const PricingCard = ({ tool }: { tool: (typeof tools)[0] }) => (
    <div className="group border border-[#1f1f1f] hover:border-[#2e2e2e] bg-[#0f0f0f] hover:bg-[#161616] p-7 flex flex-col transition-all duration-200 cursor-default">
      {/* Tool Name */}
      <div className="flex items-center justify-between mb-7">
        <div className="flex items-center gap-3">
          <div
            className="w-10 h-10 flex items-center justify-center border border-[#2e2e2e] font-display text-[22px] leading-none"
            style={{ color: tool.accentColor }}
          >
            {tool.letterMark}
          </div>
          <span className="font-display text-[24px] leading-none text-[#f0f0f0] uppercase">
            {tool.name}
          </span>
        </div>
        <span className="font-mono text-[10px] tracking-[0.15em] text-[#666] uppercase">
          {tool.category}
        </span>
      </div>

      {/* Pricing Rows */}
      <div className="flex flex-col gap-3 flex-1">
        {/* Free */}
        <div className="flex items-start justify-between py-3.5 border-t border-[#1f1f1f]">
          <span className="font-mono text-[11px] tracking-[0.1em] text-[#666] uppercase">
            Free
          </span>
          <span className="font-sans text-[12px] text-[#888] text-right max-w-[60%]">
            {tool.freeTier}
          </span>
        </div>

        {/* Pro */}
        <div className="flex items-start justify-between py-3.5 border-t border-[#1f1f1f]">
          <span className="font-mono text-[11px] tracking-[0.1em] text-[#3B82F6] uppercase">
            Pro
          </span>
          <span className="font-display text-[26px] leading-none text-[#f0f0f0]">
            {tool.proPriceMonthly}
          </span>
        </div>

        {/* Business */}
        <div className="flex items-start justify-between py-3.5 border-t border-[#1f1f1f]">
          <span className="font-mono text-[11px] tracking-[0.1em] text-[#666] uppercase">
            Business
          </span>
          <span className="font-sans text-[12px] text-[#888] text-right max-w-[60%]">
            {tool.businessPrice}
          </span>
        </div>
      </div>

      {/* Verdict */}
      <div className="border-t border-[#1f1f1f] pt-5 mt-4">
        <div className="font-mono text-[11px] tracking-[0.1em] text-[#666] uppercase mb-2.5">
          Verdict
        </div>
        <p className="font-sans text-[13px] leading-[1.6] text-[#888]">{tool.verdict}</p>
      </div>
    </div>
  );

  return (
    <section id="pricing" className="border-b border-[#1f1f1f]">
      {/* Section Header */}
      <div className="max-w-[1280px] mx-auto px-6 py-16 border-b border-[#1f1f1f]">
        <div className="flex items-center gap-4">
          <span className="font-mono text-[11px] tracking-[0.25em] text-[#3B82F6] uppercase">
            04 / Pricing
          </span>
          <div className="h-[1px] flex-1 bg-[#1f1f1f]" />
        </div>
        <h2 className="font-display text-[clamp(48px,6vw,96px)] leading-[0.95] text-[#f0f0f0] uppercase mt-6">
          Pricing
          <br />
          <span className="text-[#3B82F6]">Breakdown</span>
        </h2>
        <p className="font-sans text-[14px] text-[#666] mt-4 tracking-[0.05em]">
          All pricing reflects Q1 2026 published rates
        </p>
      </div>

      <div className="max-w-[1280px] mx-auto px-6 py-10">
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-4">
          {row1.map((tool) => (
            <PricingCard key={tool.id} tool={tool} />
          ))}
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 mb-4">
          {row2.map((tool) => (
            <PricingCard key={tool.id} tool={tool} />
          ))}
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
          {row3.map((tool) => (
            <PricingCard key={tool.id} tool={tool} />
          ))}
        </div>
      </div>
    </section>
  );
}
