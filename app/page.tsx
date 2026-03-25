import Navbar from '@/components/Navbar';
import Hero from '@/components/Hero';
import WhySection from '@/components/WhySection';
import ToolsGrid from '@/components/ToolsGrid';
import ComparisonTable from '@/components/ComparisonTable';
import PricingSection from '@/components/PricingSection';
import SuperBuilderSection from '@/components/SuperBuilderSection';
import Footer from '@/components/Footer';

export default function Home() {
  return (
    <main className="min-h-screen bg-[#080808]">
      <Navbar />
      <Hero />
      <WhySection />
      <ToolsGrid />
      <ComparisonTable />
      <PricingSection />
      <SuperBuilderSection />
      <Footer />
    </main>
  );
}
