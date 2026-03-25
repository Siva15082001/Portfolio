import { reasons } from '@/data/reasons';

export default function WhySection() {
  return (
    <section id="why-vibe" className="border-b border-[#1f1f1f]">
      {/* Section Header */}
      <div className="max-w-[1280px] mx-auto px-6 py-16 border-b border-[#1f1f1f]">
        <div className="flex items-center gap-4">
          <span className="font-mono text-[11px] tracking-[0.25em] text-[#3B82F6] uppercase">
            01 / Why
          </span>
          <div className="h-[1px] flex-1 bg-[#1f1f1f]" />
        </div>
        <h2 className="font-display text-[clamp(48px,6vw,96px)] leading-[0.95] text-[#f0f0f0] uppercase mt-6">
          Why Motive Needs
          <br />
          <span className="text-[#3B82F6]">Vibe Coding</span>
        </h2>
      </div>

      {/* Cards Row */}
      <div className="max-w-[1280px] mx-auto px-6 py-0">
        <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-5">
          {reasons.map((reason, index) => (
            <div
              key={reason.number}
              className={`why-card group relative px-7 py-12 border-b border-[#1f1f1f] lg:border-b-0 ${
                index < reasons.length - 1 ? 'lg:border-r border-[#1f1f1f]' : ''
              } cursor-default`}
            >
              {/* Hover bottom bar */}
              <div className="why-card-bar absolute bottom-0 left-0 h-[3px] bg-[#3B82F6] w-0 group-hover:w-full transition-all duration-500 ease-out" />

              {/* Number */}
              <div className="font-display text-[80px] leading-none text-[#1f1f1f] group-hover:text-[#2e2e2e] transition-colors duration-300 mb-7 select-none">
                {reason.number}
              </div>

              {/* Title */}
              <h3 className="font-sans text-[14px] font-semibold tracking-[0.1em] text-[#f0f0f0] uppercase mb-3.5">
                {reason.title}
              </h3>

              {/* Description */}
              <p className="font-sans text-[14px] leading-[1.7] text-[#888]">
                {reason.description}
              </p>
            </div>
          ))}
        </div>
      </div>
    </section>
  );
}
