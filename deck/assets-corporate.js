/**
 * Illustration and icon assets for the corporate "Parallelism (in Grammar)" deck.
 *
 * Line-art illustrations are generated as SVG and rasterised with sharp, so the
 * deck carries real drawings rather than clip art. Two of them deliberately echo
 * the source presentation's own images: parallel railway tracks and a balance.
 */
const sharp = require("sharp");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const Fa = require("react-icons/fa6");

/** Corporate palette. Red means "not parallel", green means "parallel". */
const C = {
  navy: "16233A", // dominant: headings, title and closing slides
  navySoft: "2A3C5E",
  slate: "3B455C", // body copy
  muted: "77839A",
  hairline: "DFE5EE",
  panel: "F6F8FB",
  white: "FFFFFF",
  ochre: "B4832B", // accent: eyebrows, rule numbers, emphasis
  ochreTint: "FBF4E6",
  red: "B3261E", // NOT PARALLEL
  redTint: "FBEDEB",
  redLine: "E7C3BF",
  green: "1E7A4F", // PARALLEL
  greenTint: "EAF4EF",
  greenLine: "BFDCCB",
};

const asData = (buf) => "image/png;base64," + buf.toString("base64");
const png = (svg, w, h) =>
  sharp(Buffer.from(svg)).resize(w, h, { fit: "fill" }).png().toBuffer();

/**
 * Parallel railway track, drawn flat and to scale — the source deck's opening
 * metaphor, redrawn as line art.
 */
async function rails({ dark = false, w = 1200, h = 520 } = {}) {
  const ink = dark ? "#FFFFFF" : "#" + C.navy;
  const accent = dark ? "#E7C879" : "#" + C.ochre;
  const sleepers = [];
  for (let i = 0; i < 13; i++) {
    const x = 70 + i * 88;
    sleepers.push(
      `<rect x="${x}" y="170" width="34" height="180" rx="6" fill="${accent}"
         fill-opacity="0.18" stroke="${accent}" stroke-width="3"/>`
    );
  }
  return asData(
    await png(
      `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 1200 520">
        ${sleepers.join("")}
        <rect x="40" y="196" width="1120" height="20" rx="10" fill="${ink}" fill-opacity="0.9"/>
        <rect x="40" y="306" width="1120" height="20" rx="10" fill="${ink}" fill-opacity="0.9"/>
        <text x="600" y="120" text-anchor="middle" font-family="Cambria,Georgia,serif"
              font-size="46" font-style="italic" fill="${ink}" fill-opacity="0.85">
          two rails, one shape, all the way along
        </text>
        <text x="600" y="440" text-anchor="middle" font-family="Calibri,Arial" font-size="30"
              letter-spacing="6" font-weight="700" fill="${accent}">PARALLEL</text>
      </svg>`,
      w,
      h
    )
  );
}

/** Balance beam — the source deck's illustration of "same pattern of words". */
async function balance({ w = 940, h = 620 } = {}) {
  const ink = "#" + C.navy;
  const gold = "#" + C.ochre;
  const green = "#" + C.green;
  const block = (x, y, colour) =>
    `<rect x="${x}" y="${y}" width="150" height="38" rx="19" fill="${colour}" fill-opacity="0.18"
       stroke="${colour}" stroke-width="3"/>`;
  return asData(
    await png(
      `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 940 620">
        <line x1="140" y1="250" x2="800" y2="250" stroke="${ink}" stroke-width="8" stroke-linecap="round"/>
        <path d="M470 250 L405 470 L535 470 Z" fill="none" stroke="${ink}" stroke-width="8" stroke-linejoin="round"/>
        <line x1="345" y1="470" x2="595" y2="470" stroke="${ink}" stroke-width="8" stroke-linecap="round"/>
        <line x1="215" y1="250" x2="215" y2="322" stroke="${ink}" stroke-width="5"/>
        <line x1="725" y1="250" x2="725" y2="322" stroke="${ink}" stroke-width="5"/>
        ${block(140, 322, green)}${block(140, 374, green)}
        ${block(650, 322, green)}${block(650, 374, green)}
        <circle cx="470" cy="250" r="16" fill="${ink}"/>
        <text x="215" y="185" text-anchor="middle" font-family="Cambria,Georgia,serif"
              font-size="40" fill="${ink}">same pattern</text>
        <text x="725" y="185" text-anchor="middle" font-family="Cambria,Georgia,serif"
              font-size="40" fill="${ink}">same pattern</text>
        <text x="470" y="560" text-anchor="middle" font-family="Calibri,Arial" font-size="32"
              letter-spacing="5" font-weight="700" fill="${gold}">BALANCE</text>
      </svg>`,
      w,
      h
    )
  );
}

/** Two matched blocks joined by a conjunction — the four-contexts overview. */
async function bridge({ w = 1200, h = 300 } = {}) {
  const ink = "#" + C.navy;
  const gold = "#" + C.ochre;
  const green = "#" + C.green;
  return asData(
    await png(
      `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 1200 300">
        <rect x="40" y="70" width="360" height="110" rx="16" fill="${green}" fill-opacity="0.12"
              stroke="${green}" stroke-width="4"/>
        <rect x="800" y="70" width="360" height="110" rx="16" fill="${green}" fill-opacity="0.12"
              stroke="${green}" stroke-width="4"/>
        <text x="220" y="142" text-anchor="middle" font-family="Cambria,Georgia,serif"
              font-size="44" fill="${ink}">element A</text>
        <text x="980" y="142" text-anchor="middle" font-family="Cambria,Georgia,serif"
              font-size="44" fill="${ink}">element B</text>
        <line x1="400" y1="125" x2="800" y2="125" stroke="${gold}" stroke-width="5"
              stroke-dasharray="14 12"/>
        <circle cx="600" cy="125" r="56" fill="#FFFFFF" stroke="${gold}" stroke-width="5"/>
        <text x="600" y="140" text-anchor="middle" font-family="Calibri,Arial" font-size="32"
              font-weight="700" fill="${gold}">join</text>
        <text x="600" y="255" text-anchor="middle" font-family="Calibri,Arial" font-size="30"
              fill="${ink}" fill-opacity="0.75">
          and · or · but · not only…but also · than · as · ,
        </text>
      </svg>`,
      w,
      h
    )
  );
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
  return asData(await png(markup, px, px));
}

module.exports = { C, rails, balance, bridge, icon };
