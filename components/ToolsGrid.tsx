import { tools } from '@/data/tools';
import type { ToolCategory } from '@/data/tools';

const categoryColors: Record<ToolCategory, string> = {
  IDE: 'text-[#3b82f6] border-[#3b82f6]/30 bg-[#3b82f6]/10',
  Builder: 'text-[#22c55e] border-[#22c55e]/30 bg-[#22c55e]/10',
  Prototyping: 'text-[#eab308] border-[#eab308]/30 bg-[#eab308]/10',
  Agent: 'text-[#FF6200] border-[#FF6200]/30 bg-[#FF6200]/10',
  'UI Generator': 'text-[#a855f7] border-[#a855f7]/30 bg-[#a855f7]/10',
  Terminal: 'text-[#f43f5e] border-[#f43f5e]/30 bg-[#f43f5e]/10',
};

export default function ToolsGrid() {
  return (
    <section id="tools" className="border-b border-[#1f1f1f]">
      {/* Section Header */}
      <div className="max-w-[1280px] mx-auto px-6 py-12 border-b border-[#1f1f1f]">
        <div className="flex items-center gap-4">
          <span className="font-mono text-[10px] tracking-[0.25em] text-[#FF6200] uppercase">
            02 / Tools
          </span>
          <div className="h-[1px] flex-1 bg-[#1f1f1f]" />
        </div>
        <h2 className="font-display text-[clamp(36px,5vw,72px)] leading-none text-[#f0f0f0] uppercase mt-4">
          The 8 Tools
          <br />
          <span className="text-[#FF6200]">That Matter</span>
        </h2>
      </div>

      {/* Grid */}
      <div className="max-w-[1280px] mx-auto px-6 py-0">
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4">
          {tools.map((tool, index) => {
            const colBorder = index % 4 !== 3 ? 'lg:border-r' : '';
            const rowBorder =
              index < tools.length - 4 ? 'border-b' : 'border-b lg:border-b-0';
            return (
              <div
                key={tool.id}
                className={`tool-card group relative flex flex-col ${rowBorder} ${colBorder} border-[#1f1f1f] p-6 cursor-default overflow-hidden`}
              >
                {/* Top accent bar on hover */}
                <div
                  className="tool-card-bar absolute top-0 left-0 right-0 h-[2px] scale-x-0 group-hover:scale-x-100 transition-transform duration-300 ease-out origin-left"
                  style={{ backgroundColor: tool.accentColor }}
                />

                {/* Hover background lift */}
                <div className="absolute inset-0 bg-[#0f0f0f] opacity-0 group-hover:opacity-100 transition-opacity duration-300" />

                <div className="relative z-10 flex flex-col h-full">
                  {/* Logo + Category */}
                  <div className="flex items-start justify-between mb-5">
                    <div
                      className="w-10 h-10 flex items-center justify-center border border-[#2e2e2e] font-display text-[22px] leading-none"
                      style={{ color: tool.accentColor }}
                    >
                      {tool.letterMark}
                    </div>
                    <span
                      className={`font-mono text-[9px] tracking-[0.15em] uppercase border px-2 py-1 ${
                        categoryColors[tool.category]
                      }`}
                    >
                      {tool.category}
                    </span>
                  </div>

                  {/* Name */}
                  <h3 className="font-display text-[32px] leading-none text-[#f0f0f0] mb-2 uppercase">
                    {tool.name}
                  </h3>

                  {/* Tagline */}
                  <p className="font-mono text-[11px] leading-[1.6] text-[#666] mb-5">
                    {tool.tagline}
                  </p>

                  {/* Features */}
                  <ul className="flex flex-col gap-2 mb-6 flex-1">
                    {tool.features.map((feature, i) => (
                      <li key={i} className="flex items-start gap-2">
                        <span
                          className="mt-[5px] w-[4px] h-[4px] rounded-full flex-shrink-0"
                          style={{ backgroundColor: tool.accentColor }}
                        />
                        <span className="font-mono text-[11px] leading-[1.5] text-[#888]">
                          {feature}
                        </span>
                      </li>
                    ))}
                  </ul>

                  {/* Footer */}
                  <div className="border-t border-[#1f1f1f] pt-4 flex items-center justify-between">
                    <div>
                      <div className="font-mono text-[10px] tracking-[0.1em] text-[#FF6200] uppercase">
                        From {tool.proPriceMonthly}
                      </div>
                      <div className="font-mono text-[10px] text-[#444] mt-1">
                        {tool.bestFor}
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </section>
  );
}
