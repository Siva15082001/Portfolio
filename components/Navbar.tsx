'use client';

import { useEffect, useState } from 'react';

export default function Navbar() {
  const [scrollProgress, setScrollProgress] = useState(0);
  const [scrolled, setScrolled] = useState(false);

  useEffect(() => {
    const handleScroll = () => {
      const totalHeight =
        document.documentElement.scrollHeight - document.documentElement.clientHeight;
      const progress = (window.scrollY / totalHeight) * 100;
      setScrollProgress(progress);
      setScrolled(window.scrollY > 20);
    };
    window.addEventListener('scroll', handleScroll, { passive: true });
    return () => window.removeEventListener('scroll', handleScroll);
  }, []);

  return (
    <nav
      className={`fixed top-0 left-0 right-0 z-50 transition-all duration-300 ${
        scrolled ? 'bg-[#080808]/90 backdrop-blur-md border-b border-[#1f1f1f]' : 'bg-transparent'
      }`}
    >
      {/* Scroll Progress Bar */}
      <div
        className="absolute bottom-0 left-0 h-[2px] bg-[#FF6200] transition-all duration-100 ease-out"
        style={{ width: `${scrollProgress}%` }}
      />

      <div className="max-w-[1280px] mx-auto px-6 h-14 flex items-center justify-between">
        {/* Logo */}
        <div className="flex items-center gap-3">
          <div className="w-7 h-7 bg-[#FF6200] flex items-center justify-center">
            <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
              <path
                d="M2 3L7 1L12 3V7C12 10 7 13 7 13C7 13 2 10 2 7V3Z"
                fill="white"
                opacity="0.9"
              />
              <path d="M5 6.5L7 8.5L11 4.5" stroke="#FF6200" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          </div>
          <span className="font-mono text-[11px] tracking-[0.2em] text-[#f0f0f0] uppercase">
            Motive Eng
          </span>
        </div>

        {/* Nav Links */}
        <div className="hidden md:flex items-center gap-8">
          {['Why Vibe', 'Tools', 'Compare', 'Pricing', 'Use Cases'].map((link) => (
            <a
              key={link}
              href={`#${link.toLowerCase().replace(/ /g, '-')}`}
              className="font-mono text-[11px] tracking-[0.15em] text-[#666] hover:text-[#FF6200] transition-colors duration-200 uppercase"
            >
              {link}
            </a>
          ))}
        </div>

        {/* Right Tag */}
        <div className="border border-[#FF6200]/40 px-3 py-1">
          <span className="font-mono text-[10px] tracking-[0.2em] text-[#FF6200] uppercase">
            Q2 2026
          </span>
        </div>
      </div>
    </nav>
  );
}
