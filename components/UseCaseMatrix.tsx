import { useCases } from '@/data/usecases';

const toolColors: Record<string, string> = {
  Lovable: '#EC4899',
  Bolt: '#0EA5E9',
  v0: '#a855f7',
  Cursor: '#7C3AED',
  Windsurf: '#06B6D4',
  Replit: '#F26207',
  'GitHub Copilot': '#22C55E',
  'Claude Code': '#D97706',
};

export default function UseCaseMatrix() {
  return (
    <section id="use-cases" className="border-b border-[#1f1f1f]">
      {/* Section Header */}
      <div className="max-w-[1280px] mx-auto px-6 py-12 border-b border-[#1f1f1f]">
        <div className="flex items-center gap-4">
          <span className="font-mono text-[10px] tracking-[0.25em] text-[#FF6200] uppercase">
            05 / Use Cases
          </span>
          <div className="h-[1px] flex-1 bg-[#1f1f1f]" />
        </div>
        <h2 className="font-display text-[clamp(36px,5vw,72px)] leading-none text-[#f0f0f0] uppercase mt-4">
          Who Uses
          <br />
          <span className="text-[#FF6200]">What at Motive</span>
        </h2>
      </div>

      <div className="max-w-[1280px] mx-auto px-6 py-10">
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-0">
          {useCases.map((useCase, index) => {
            const borders = [
              'border-r border-b', // 0
              'border-r border-b', // 1
              'border-b',          // 2
              'border-r',          // 3
              'border-r',          // 4
              '',                  // 5
            ];
            return (
              <div
                key={useCase.role}
                className={`group p-8 border-[#1f1f1f] hover:bg-[#0f0f0f] transition-colors duration-200 cursor-default ${borders[index]}`}
              >
                {/* Role */}
                <div className="font-mono text-[10px] tracking-[0.25em] text-[#444] uppercase mb-4">
                  Motive Role
                </div>
                <h3 className="font-display text-[28px] leading-none text-[#f0f0f0] uppercase mb-5">
                  {useCase.role}
                </h3>

                {/* Tool Tags */}
                <div className="flex flex-wrap gap-2 mb-5">
                  {useCase.tools.map((toolName) => (
                    <span
                      key={toolName}
                      className="font-mono text-[10px] tracking-[0.1em] uppercase px-2 py-1 border"
                      style={{
                        color: toolColors[toolName] || '#888',
                        borderColor: `${toolColors[toolName]}30` || '#1f1f1f',
                        backgroundColor: `${toolColors[toolName]}10` || 'transparent',
                      }}
                    >
                      {toolName}
                    </span>
                  ))}
                </div>

                {/* Description */}
                <p className="font-mono text-[12px] leading-[1.7] text-[#666]">
                  {useCase.description}
                </p>
              </div>
            );
          })}
        </div>
      </div>
    </section>
  );
}
