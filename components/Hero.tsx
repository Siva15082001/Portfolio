export default function Hero() {
  const stats = [
    { value: '8', label: 'Tools Analyzed' },
    { value: '41%', label: 'Code is AI-Generated Globally' },
    { value: '$9.9B', label: 'Cursor Valuation' },
    { value: '2025', label: 'Collins Word of the Year: Vibe Coding' },
  ];

  return (
    <section
      id="hero"
      className="relative min-h-screen flex flex-col justify-center pt-14 overflow-hidden"
    >
      {/* Grid Background */}
      <div className="absolute inset-0 grid-bg pointer-events-none" />

      {/* Blue Radial Glow */}
      <div
        className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[800px] h-[600px] pointer-events-none"
        style={{
          background:
            'radial-gradient(ellipse at center, rgba(59,130,246,0.08) 0%, rgba(59,130,246,0.02) 50%, transparent 70%)',
        }}
      />

      <div className="max-w-[1280px] mx-auto px-6 w-full">
        {/* Eyebrow */}
        <div className="hero-fade-1 mb-6">
          <span className="font-mono text-[13px] tracking-[0.25em] text-[#3B82F6] uppercase">
            Product Team Reference
          </span>
        </div>

        {/* Main Headline */}
        <h1 className="hero-fade-2 font-display text-[clamp(80px,11vw,180px)] leading-[0.9] tracking-[-0.02em] text-[#f0f0f0] uppercase mb-5">
          Vibe Coding
          <br />
          <span className="text-[#3B82F6]">Landscape</span>
        </h1>

        {/* Subheading */}
        <p className="hero-fade-3 font-sans text-[16px] tracking-[0.05em] text-[#888] mb-16 max-w-lg">
          Motive Technologies Inc — Q2 2026 Tool Analysis
        </p>

        {/* Stat Pills */}
        <div className="hero-fade-4 flex flex-wrap gap-4">
          {stats.map((stat) => (
            <div
              key={stat.label}
              className="border border-[#1f1f1f] bg-[#0f0f0f] px-6 py-4 flex flex-col gap-1.5"
            >
              <span className="font-display text-[36px] leading-none text-[#f0f0f0]">
                {stat.value}
              </span>
              <span className="font-sans text-[12px] tracking-[0.1em] text-[#666] uppercase max-w-[180px]">
                {stat.label}
              </span>
            </div>
          ))}
        </div>
      </div>

      {/* Bottom border */}
      <div className="absolute bottom-0 left-0 right-0 h-[1px] bg-[#1f1f1f]" />
    </section>
  );
}
