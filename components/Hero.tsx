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

      {/* Orange Radial Glow */}
      <div
        className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[800px] h-[600px] pointer-events-none"
        style={{
          background:
            'radial-gradient(ellipse at center, rgba(255,98,0,0.06) 0%, rgba(255,98,0,0.01) 50%, transparent 70%)',
        }}
      />

      <div className="max-w-[1280px] mx-auto px-6 w-full">
        {/* Eyebrow */}
        <div className="hero-fade-1 mb-6">
          <span className="font-mono text-[11px] tracking-[0.25em] text-[#FF6200] uppercase">
            Internal Engineering Reference
          </span>
        </div>

        {/* Main Headline */}
        <h1 className="hero-fade-2 font-display text-[clamp(72px,10vw,160px)] leading-[0.9] tracking-[-0.02em] text-[#f0f0f0] uppercase mb-4">
          Vibe Coding
          <br />
          <span className="text-[#FF6200]">Landscape</span>
        </h1>

        {/* Subheading */}
        <p className="hero-fade-3 font-mono text-[14px] tracking-[0.1em] text-[#666] uppercase mb-16 max-w-lg">
          Motive — Q2 2026 Tool Analysis
        </p>

        {/* Stat Pills */}
        <div className="hero-fade-4 flex flex-wrap gap-3">
          {stats.map((stat) => (
            <div
              key={stat.label}
              className="border border-[#1f1f1f] bg-[#0f0f0f] px-5 py-3 flex flex-col gap-1"
            >
              <span className="font-display text-[28px] leading-none text-[#f0f0f0]">
                {stat.value}
              </span>
              <span className="font-mono text-[10px] tracking-[0.15em] text-[#666] uppercase max-w-[160px]">
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
