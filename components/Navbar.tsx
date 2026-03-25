'use client';

import { useEffect, useState } from 'react';
import Image from 'next/image';

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
        className="absolute bottom-0 left-0 h-[2px] bg-[#3B82F6] transition-all duration-100 ease-out"
        style={{ width: `${scrollProgress}%` }}
      />

      <div className="max-w-[1280px] mx-auto px-6 h-16 flex items-center justify-between">
        {/* Logo */}
        <div className="flex items-center">
          <Image
            src="/motive-logo.png"
            alt="Motive"
            width={120}
            height={32}
            className="h-8 w-auto"
            priority
          />
        </div>

        {/* Nav Links */}
        <div className="hidden md:flex items-center gap-8">
          {['Why Vibe', 'Tools', 'Compare', 'Pricing', 'Recommendation'].map((link) => (
            <a
              key={link}
              href={`#${link.toLowerCase().replace(/ /g, '-')}`}
              className="font-mono text-[12px] tracking-[0.15em] text-[#666] hover:text-[#3B82F6] transition-colors duration-200 uppercase"
            >
              {link}
            </a>
          ))}
        </div>

        {/* Right Tag */}
        <div className="border border-[#3B82F6]/40 px-3 py-1.5">
          <span className="font-mono text-[11px] tracking-[0.2em] text-[#3B82F6] uppercase">
            Q2 2026
          </span>
        </div>
      </div>
    </nav>
  );
}
