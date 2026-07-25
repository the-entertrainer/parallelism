# Parallelism

**Interactive training web app** for instructional designers and content developers on the art and science of parallel structure.

## Experience

- 3D dynamic background (Three.js parallel lines + particles)
- Glassmorphism UI with depth, blur, and subtle morph-style transitions
- Icebreaker hook → deep dive → practical examples → interactive practice → summary
- Keyboard, mouse, and touch navigation
- Full accessibility: reduced-motion support, ARIA, focus management, semantic structure
- Mobile-first responsive design

## Run locally

Because this is pure web (no build step required), simply open `index.html` in a modern browser, or serve the folder:

```bash
npx serve .
# or
python -m http.server 5173
```

Then visit the local URL.

Libraries load from CDN (Three.js). Internet required for first load of the modules and Google Fonts.

## Design notes

- Design tokens defined in CSS custom properties
- Glass panels use `backdrop-filter` + layered borders/highlights
- Transitions respect `prefers-reduced-motion`
- Progress rail + dots for orientation
- Practice questions give immediate feedback

## Theme

Parallelism is treated as both a writing craft and a broader design principle that extends to visuals, motion, and interaction patterns in learning experiences.

Built with care for clarity, delight, and accessibility.
