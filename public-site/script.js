const observer = new IntersectionObserver((entries) => {
  entries.forEach((entry) => {
    if (entry.isIntersecting) entry.target.classList.add('is-visible');
  });
}, { threshold: 0.16 });

document.querySelectorAll('.section, .product-card').forEach((el) => {
  el.classList.add('reveal');
  observer.observe(el);
});