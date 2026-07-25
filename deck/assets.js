/**
 * Asset generation for the Parallelism training deck.
 * Produces the gradient backgrounds (pptxgenjs cannot emit gradient fills)
 * and rasterised react-icons, all as base64 PNG data ready for addImage().
 */
const sharp = require("sharp");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const Fa = require("react-icons/fa6");

const W = 1920;
const H = 1080;

/** Deck palette. One tone means "breaks", one means "holds". */
const C = {
  ink: "05070F", // deepest base tone
  ink2: "0B1020",
  text: "F4F6FF",
  textDim: "A7AFC9",
  textFaint: "6F779A",
  iris: "8A7BFF", // structural / neutral accent (the "current item" colour)
  irisDeep: "3A2CA8",
  breaks: "FF4F6E", // rose — mismatched form
  holds: "2FE0A5", // mint — parallel form
  white: "FFFFFF",
};

function svgToPng(svg, w, h) {
  return sharp(Buffer.from(svg)).resize(w, h, { fit: "fill" }).png().toBuffer();
}

const asData = (buf) => "image/png;base64," + buf.toString("base64");

/** Dark mesh-gradient background. `hero` pushes the glows brighter. */
async function background(hero) {
  const a = hero ? 0.85 : 0.5; // glow strength multiplier
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}">
    <defs>
      <linearGradient id="base" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0%" stop-color="#0A0F20"/>
        <stop offset="55%" stop-color="#06090F"/>
        <stop offset="100%" stop-color="#0B0A18"/>
      </linearGradient>
      <radialGradient id="g1" cx="50%" cy="50%" r="50%">
        <stop offset="0%" stop-color="#6E5CFF" stop-opacity="${0.5 * a}"/>
        <stop offset="60%" stop-color="#6E5CFF" stop-opacity="${0.12 * a}"/>
        <stop offset="100%" stop-color="#6E5CFF" stop-opacity="0"/>
      </radialGradient>
      <radialGradient id="g2" cx="50%" cy="50%" r="50%">
        <stop offset="0%" stop-color="#2FE0A5" stop-opacity="${0.26 * a}"/>
        <stop offset="100%" stop-color="#2FE0A5" stop-opacity="0"/>
      </radialGradient>
      <radialGradient id="g3" cx="50%" cy="50%" r="50%">
        <stop offset="0%" stop-color="#FF4F6E" stop-opacity="${0.24 * a}"/>
        <stop offset="100%" stop-color="#FF4F6E" stop-opacity="0"/>
      </radialGradient>
    </defs>
    <rect width="${W}" height="${H}" fill="url(#base)"/>
    <ellipse cx="${W * 0.18}" cy="${H * 0.12}" rx="${W * 0.42}" ry="${H * 0.5}" fill="url(#g1)"/>
    <ellipse cx="${W * 0.92}" cy="${H * 0.18}" rx="${W * 0.3}" ry="${H * 0.42}" fill="url(#g2)"/>
    <ellipse cx="${W * 0.72}" cy="${H * 1.02}" rx="${W * 0.4}" ry="${H * 0.45}" fill="url(#g3)"/>
    <ellipse cx="${W * 0.04}" cy="${H * 0.95}" rx="${W * 0.22}" ry="${H * 0.3}" fill="url(#g1)"/>
  </svg>`;
  return asData(await svgToPng(svg, W, H));
}

/** Soft white specular blob — the highlight that sells the glass. */
async function sheen(opacity = 0.5) {
  const s = 900;
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${s}" height="${s}">
    <defs><radialGradient id="s" cx="50%" cy="50%" r="50%">
      <stop offset="0%" stop-color="#FFFFFF" stop-opacity="${opacity}"/>
      <stop offset="45%" stop-color="#FFFFFF" stop-opacity="${opacity * 0.28}"/>
      <stop offset="100%" stop-color="#FFFFFF" stop-opacity="0"/>
    </radialGradient></defs>
    <circle cx="${s / 2}" cy="${s / 2}" r="${s / 2}" fill="url(#s)"/>
  </svg>`;
  return asData(await svgToPng(svg, s, s));
}

/** Rasterise a react-icons glyph in a given hex colour. */
async function icon(name, hex, px = 256) {
  const Comp = Fa[name];
  if (!Comp) throw new Error("Unknown icon: " + name);
  let markup = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: "#" + hex, size: px })
  );
  if (!markup.includes("xmlns")) {
    markup = markup.replace("<svg", '<svg xmlns="http://www.w3.org/2000/svg"');
  }
  return asData(await svgToPng(markup, px, px));
}

module.exports = { C, background, sheen, icon, asData };
