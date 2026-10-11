/* Domainzax lightweight motion system. Native browser APIs and CSS only. */
(() => {
  const reduceMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  if (reduceMotion) return;
  const targets = document.querySelectorAll(".hero .eyebrow, .hero h1, .hero-grid, .section-head, .spotlight, .stat, .contact");
  if ("IntersectionObserver" in window) {
    const observer = new IntersectionObserver((entries, instance) => entries.forEach(entry => {
      if (!entry.isIntersecting) return;
      entry.target.classList.add("motion-visible");
      instance.unobserve(entry.target);
    }), { threshold: 0.08, rootMargin: "0px 0px -30px 0px" });
    targets.forEach((element, index) => {
      element.classList.add("motion-ready");
      element.style.setProperty("--motion-delay", Math.min(index % 5, 4) * 65 + "ms");
      observer.observe(element);
    });
  }
  if (!window.matchMedia("(hover: hover) and (pointer: fine)").matches) return;
  document.querySelectorAll(".asset, .spotlight, .contact").forEach(card => card.addEventListener("pointermove", event => {
    const rect = card.getBoundingClientRect();
    card.style.setProperty("--pointer-x", event.clientX - rect.left + "px");
    card.style.setProperty("--pointer-y", event.clientY - rect.top + "px");
  }, { passive: true }));
})();
