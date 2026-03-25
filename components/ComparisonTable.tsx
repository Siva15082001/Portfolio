import { tools } from '@/data/tools';

type CheckValue = 'yes' | 'no' | 'partial';

const Check = ({ value }: { value: CheckValue | string }) => {
  if (value === 'yes') {
    return <span className="text-[#22c55e] font-mono text-[13px]">&#10003;</span>;
  }
  if (value === 'no') {
    return <span className="text-[#444] font-mono text-[13px]">&#8212;</span>;
  }
  if (value === 'partial') {
    return <span className="text-[#eab308] font-mono text-[13px]">&#9680;</span>;
  }
  return <span className="font-mono text-[11px] text-[#888]">{value}</span>;
};

const columnHeaders = [
  'Tool',
  'Type',
  'AI Model',
  'Full-Stack',
  'GitHub Sync',
  'Free Tier',
  'Deployment',
  'Best For',
];

export default function ComparisonTable() {
  return (
    <section id="compare" className="border-b border-[#1f1f1f]">
      {/* Section Header */}
      <div className="max-w-[1280px] mx-auto px-6 py-12 border-b border-[#1f1f1f]">
        <div className="flex items-center gap-4">
          <span className="font-mono text-[10px] tracking-[0.25em] text-[#FF6200] uppercase">
            03 / Compare
          </span>
          <div className="h-[1px] flex-1 bg-[#1f1f1f]" />
        </div>
        <h2 className="font-display text-[clamp(36px,5vw,72px)] leading-none text-[#f0f0f0] uppercase mt-4">
          Side-by-Side
          <br />
          <span className="text-[#FF6200]">Comparison</span>
        </h2>
      </div>

      <div className="max-w-[1280px] mx-auto px-6 py-10 overflow-x-auto">
        <table className="w-full border-collapse min-w-[900px]">
          <thead>
            <tr className="border-b border-[#2e2e2e]">
              {columnHeaders.map((h) => (
                <th
                  key={h}
                  className="text-left font-mono text-[10px] tracking-[0.2em] text-[#444] uppercase pb-3 pr-6 first:pl-0"
                >
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {tools.map((tool, i) => (
              <tr
                key={tool.id}
                className={`border-b border-[#1f1f1f] group hover:bg-[#0f0f0f] transition-colors duration-150 ${
                  i % 2 === 0 ? '' : 'bg-[#080808]'
                }`}
              >
                {/* Tool */}
                <td className="py-4 pr-6">
                  <div className="flex items-center gap-3">
                    <div
                      className="w-6 h-6 flex items-center justify-center border border-[#2e2e2e] font-display text-[13px] flex-shrink-0"
                      style={{ color: tool.accentColor }}
                    >
                      {tool.letterMark}
                    </div>
                    <span className="font-sans text-[13px] font-semibold text-[#f0f0f0]">
                      {tool.name}
                    </span>
                  </div>
                </td>

                {/* Type */}
                <td className="py-4 pr-6">
                  <span className="font-mono text-[10px] tracking-[0.1em] text-[#666] uppercase">
                    {tool.category}
                  </span>
                </td>

                {/* AI Model */}
                <td className="py-4 pr-6 max-w-[180px]">
                  <span className="font-mono text-[11px] text-[#888] leading-[1.5] block">
                    {tool.aiModels}
                  </span>
                </td>

                {/* Full-Stack */}
                <td className="py-4 pr-6 text-center">
                  <Check value={tool.fullStack} />
                </td>

                {/* GitHub Sync */}
                <td className="py-4 pr-6 text-center">
                  <Check value={tool.githubSync} />
                </td>

                {/* Free Tier */}
                <td className="py-4 pr-6">
                  <span className="font-mono text-[11px] text-[#888]">{tool.freeTier}</span>
                </td>

                {/* Deployment */}
                <td className="py-4 pr-6 max-w-[160px]">
                  <span className="font-mono text-[11px] text-[#888]">{tool.deployment}</span>
                </td>

                {/* Best For */}
                <td className="py-4">
                  <span className="font-mono text-[11px] text-[#888] leading-[1.5] block max-w-[180px]">
                    {tool.bestFor}
                  </span>
                </td>
              </tr>
            ))}
          </tbody>
        </table>

        {/* Legend */}
        <div className="flex items-center gap-6 mt-6">
          <span className="font-mono text-[10px] tracking-[0.15em] text-[#444] uppercase">
            Legend:
          </span>
          <div className="flex items-center gap-1">
            <span className="text-[#22c55e] font-mono text-[13px]">&#10003;</span>
            <span className="font-mono text-[10px] text-[#444]">Supported</span>
          </div>
          <div className="flex items-center gap-1">
            <span className="text-[#eab308] font-mono text-[13px]">&#9680;</span>
            <span className="font-mono text-[10px] text-[#444]">Partial</span>
          </div>
          <div className="flex items-center gap-1">
            <span className="text-[#444] font-mono text-[13px]">&#8212;</span>
            <span className="font-mono text-[10px] text-[#444]">Not Supported</span>
          </div>
        </div>
      </div>
    </section>
  );
}
