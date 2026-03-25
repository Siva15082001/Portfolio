export default function Footer() {
  const currentDate = new Date().toLocaleDateString('en-US', {
    month: 'long',
    year: 'numeric',
  });

  return (
    <footer className="border-t border-[#1f1f1f]">
      <div className="max-w-[1280px] mx-auto px-6 py-12">
        {/* Top Row */}
        <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-6 pb-8 border-b border-[#1f1f1f]">
          {/* Left: Logo + Wordmark */}
          <div className="flex items-center gap-3">
            <div className="w-8 h-8 bg-[#FF6200] flex items-center justify-center">
              <svg width="16" height="16" viewBox="0 0 14 14" fill="none">
                <path
                  d="M2 3L7 1L12 3V7C12 10 7 13 7 13C7 13 2 10 2 7V3Z"
                  fill="white"
                  opacity="0.9"
                />
                <path d="M5 6.5L7 8.5L11 4.5" stroke="#FF6200" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </div>
            <div>
              <div className="font-display text-[24px] leading-none text-[#f0f0f0] uppercase">
                Motive
              </div>
              <div className="font-mono text-[10px] tracking-[0.2em] text-[#444] uppercase">
                Engineering
              </div>
            </div>
          </div>

          {/* Right: Meta */}
          <div className="flex flex-col items-start md:items-end gap-1">
            <span className="font-mono text-[11px] text-[#666]">
              Prepared by Engineering — {currentDate}
            </span>
            <span className="font-mono text-[10px] text-[#444]">
              Internal Reference Document — Not for External Distribution
            </span>
          </div>
        </div>

        {/* Bottom Row */}
        <div className="flex flex-col md:flex-row items-start md:items-center justify-between gap-4 pt-8">
          {/* Data Sources */}
          <div>
            <div className="font-mono text-[10px] tracking-[0.15em] text-[#444] uppercase mb-2">
              Data Sources
            </div>
            <p className="font-mono text-[11px] text-[#444] max-w-xl leading-[1.7]">
              cursor.com/pricing, lovable.dev/pricing, bolt.new/pricing, v0.dev/pricing,
              replit.com/pricing, windsurf.com, github.com/features/copilot, claude.com/pricing.
              All pricing reflects Q1–Q2 2026 published rates. Subject to change.
            </p>
          </div>

          {/* Stats */}
          <div className="flex items-center gap-6">
            <div className="text-center">
              <div className="font-display text-[32px] leading-none text-[#FF6200]">8</div>
              <div className="font-mono text-[10px] text-[#444] uppercase tracking-[0.1em]">
                Tools
              </div>
            </div>
            <div className="text-center">
              <div className="font-display text-[32px] leading-none text-[#FF6200]">2026</div>
              <div className="font-mono text-[10px] text-[#444] uppercase tracking-[0.1em]">
                Current
              </div>
            </div>
          </div>
        </div>
      </div>
    </footer>
  );
}
