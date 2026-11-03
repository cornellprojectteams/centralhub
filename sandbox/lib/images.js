'use strict';
// Placeholder "photos" for the fake data: small SVG drawings of a shop corner with a
// caption, so cards, emails and the inventory board have something to show. They are
// made up on the spot; no real photos ship with the sandbox.

const SCENES = {
  bench(c) {
    return `<rect x="0" y="250" width="640" height="230" fill="${c.floor}"/>`
      + `<rect x="70" y="190" width="500" height="26" rx="4" fill="${c.wood}"/>`
      + `<rect x="90" y="216" width="16" height="120" fill="${c.dark}"/><rect x="534" y="216" width="16" height="120" fill="${c.dark}"/>`
      + `<rect x="150" y="150" width="90" height="40" rx="6" fill="${c.accent}"/><rect x="270" y="132" width="40" height="58" rx="4" fill="${c.metal}"/>`
      + `<circle cx="420" cy="172" r="18" fill="${c.metal}"/><rect x="455" y="166" width="70" height="10" rx="5" fill="${c.dark}"/>`
      + `<rect x="110" y="40" width="420" height="80" rx="6" fill="${c.wall2}"/>`
      + [0, 1, 2, 3, 4, 5].map((i) => `<rect x="${130 + i * 66}" y="58" width="10" height="${30 + (i % 3) * 10}" rx="3" fill="${c.dark}"/>`).join('');
  },
  shelves(c) {
    return `<rect x="0" y="300" width="640" height="180" fill="${c.floor}"/>`
      + [0, 1, 2].map((i) => `<rect x="120" y="${90 + i * 80}" width="400" height="12" fill="${c.wood}"/>`
        + [0, 1, 2, 3].map((j) => `<rect x="${140 + j * 92}" y="${50 + i * 80}" width="${60 + (j % 2) * 14}" height="40" rx="4" fill="${j % 2 ? c.accent : c.box}"/>`).join('')).join('')
      + `<rect x="114" y="40" width="10" height="300" fill="${c.dark}"/><rect x="516" y="40" width="10" height="300" fill="${c.dark}"/>`;
  },
  floor(c) {
    return `<rect x="0" y="210" width="640" height="270" fill="${c.floor}"/>`
      + `<polygon points="0,480 230,210 250,210 60,480" fill="${c.accent}" opacity=".9"/>`
      + `<polygon points="640,480 410,210 390,210 580,480" fill="${c.accent}" opacity=".9"/>`
      + `<rect x="280" y="120" width="80" height="120" rx="6" fill="${c.wall2}"/><rect x="292" y="132" width="56" height="70" rx="4" fill="${c.metal}"/>`
      + `<rect x="470" y="250" width="70" height="90" rx="8" fill="${c.box}"/><rect x="100" y="270" width="90" height="60" rx="8" fill="${c.box}"/>`;
  },
  cabinet(c) {
    return `<rect x="0" y="330" width="640" height="150" fill="${c.floor}"/>`
      + `<rect x="200" y="60" width="240" height="280" rx="10" fill="${c.accent}"/>`
      + `<rect x="218" y="80" width="98" height="240" rx="6" fill="${c.wall2}" opacity=".35"/><rect x="324" y="80" width="98" height="240" rx="6" fill="${c.wall2}" opacity=".35"/>`
      + `<rect x="300" y="190" width="10" height="40" rx="4" fill="${c.dark}"/><rect x="330" y="190" width="10" height="40" rx="4" fill="${c.dark}"/>`
      + `<polygon points="320,110 350,160 290,160" fill="${c.box}"/><text x="320" y="152" text-anchor="middle" font-family="Arial,sans-serif" font-size="30" font-weight="700" fill="${c.dark}">!</text>`;
  },
  tool(c) {
    return `<rect x="0" y="0" width="640" height="480" fill="${c.wall2}"/>`
      + `<g fill="${c.wall}">${Array.from({ length: 12 }, (_, i) => Array.from({ length: 9 }, (_, j) => `<circle cx="${30 + i * 53}" cy="${30 + j * 53}" r="4"/>`).join('')).join('')}</g>`
      + `<rect x="190" y="150" width="210" height="110" rx="20" fill="${c.accent}"/><rect x="380" y="180" width="90" height="40" rx="10" fill="${c.metal}"/>`
      + `<rect x="230" y="250" width="70" height="140" rx="16" fill="${c.dark}"/><rect x="215" y="380" width="100" height="36" rx="10" fill="${c.box}"/>`;
  },
};

const PALETTES = [
  { wall: '#d9d4c9', wall2: '#ece8df', floor: '#9a948a', wood: '#b07a45', dark: '#3b3a38', metal: '#8c99a6', accent: '#c0392b', box: '#e0a020' },
  { wall: '#cfd8dc', wall2: '#e8eef0', floor: '#7d8790', wood: '#a0703f', dark: '#2f3a42', metal: '#9fb1bd', accent: '#0d9488', box: '#f0c050' },
  { wall: '#dcd3e6', wall2: '#efe9f5', floor: '#8e8597', wood: '#9c6b3d', dark: '#352f3d', metal: '#a69fb3', accent: '#7c3aed', box: '#f2b84b' },
  { wall: '#d6dfd1', wall2: '#ecf1e9', floor: '#86917f', wood: '#a87747', dark: '#2f3a2e', metal: '#9aa898', accent: '#1f9d5b', box: '#e8a33c' },
  { wall: '#e0d6cc', wall2: '#f3ece4', floor: '#9b8f84', wood: '#b5804b', dark: '#3d332c', metal: '#a59a91', accent: '#2563c9', box: '#efb34a' },
];

function hash(s) { let h = 2166136261; for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); } return h >>> 0; }
const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

// photo('After: tool crib shelving', 'shelves') -> a Buffer holding an SVG image.
function photo(caption, scene) {
  const h = hash(caption);
  const c = PALETTES[h % PALETTES.length];
  const draw = SCENES[scene] || SCENES[Object.keys(SCENES)[h % Object.keys(SCENES).length]];
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="640" height="480" viewBox="0 0 640 480">`
    + `<rect width="640" height="480" fill="${c.wall}"/>${draw(c)}`
    + `<rect x="0" y="420" width="640" height="60" fill="#000" opacity=".55"/>`
    + `<text x="24" y="458" font-family="Arial,Helvetica,sans-serif" font-size="22" font-weight="700" fill="#fff">${esc(caption).slice(0, 52)}</text>`
    + `<text x="616" y="458" text-anchor="end" font-family="Arial,Helvetica,sans-serif" font-size="14" font-weight="700" fill="#f0c050">SANDBOX PHOTO</text>`
    + `</svg>`;
  return Buffer.from(svg, 'utf8');
}

module.exports = { photo, SCENES: Object.keys(SCENES) };
