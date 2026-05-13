import { ShieldCheck } from "lucide-react";

export function PrivacyBadge() {
  return (
    <div className="inline-flex items-center gap-2 rounded-full border border-primary/20 bg-primary/5 px-4 py-2 text-sm text-primary backdrop-blur-sm">
      <ShieldCheck className="h-4 w-4" />
      <span>Your files are processed completely on your device and are never uploaded anywhere.</span>
    </div>
  );
}
