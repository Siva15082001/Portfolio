export default function Footer() {
  const currentDate = new Date().toLocaleDateString('en-US', {
    month: 'long',
    year: 'numeric',
  });

  return (
    <footer className="border-t border-[#1f1f1f]">
      <div className="max-w-[1280px] mx-auto px-6 py-16">
        {/* Top Row */}
        <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-8 pb-10 border-b border-[#1f1f1f]">
          {/* Left: Logo + Wordmark */}
          <div className="flex items-center gap-3">
            <div>
              <img
                src="/motive-logo.png"
                alt="Motive Technologies"
                className="h-10 w-auto mb-1"
              />
              <div className="font-sans text-[13px] tracking-[0.05em] text-[#666]">
                Technologies Inc
              </div>
            </div>
          </div>

          {/* Right: Meta */}
          <div className="flex flex-col items-start md:items-end gap-1.5">
            <span className="font-sans text-[13px] text-[#888]">
              Prepared by Product Team — {currentDate}
            </span>
            <span className="font-sans text-[12px] text-[#666]">
              Internal Reference Document — Not for External Distribution
            </span>
          </div>
        </div>

        {/* Bottom Row */}
        <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-6 pt-10">
          {/* Data Sources */}
          <div>
            <div className="font-mono text-[11px] tracking-[0.15em] text-[#666] uppercase mb-3">
              Data Sources
            </div>
            <p className="font-sans text-[13px] text-[#666] max-w-xl leading-[1.7]">
              cursor.com/pricing, lovable.dev/pricing, bolt.new/pricing, v0.dev/pricing,
              replit.com/pricing, windsurf.com, github.com/features/copilot, claude.com/pricing.
              All pricing reflects Q1–Q2 2026 published rates. Subject to change.
            </p>
          </div>

          {/* Stats */}
          <div className="flex items-center gap-8">
            <div className="text-center">
              <div className="font-display text-[40px] leading-none text-[#3B82F6]">8</div>
              <div className="font-mono text-[11px] text-[#666] uppercase tracking-[0.1em] mt-1">
                Tools
              </div>
            </div>
            <div className="text-center">
              <div className="font-display text-[40px] leading-none text-[#3B82F6]">2026</div>
              <div className="font-mono text-[11px] text-[#666] uppercase tracking-[0.1em] mt-1">
                Current
              </div>
            </div>
          </div>
        </div>
      </div>
    </footer>
  );
}
