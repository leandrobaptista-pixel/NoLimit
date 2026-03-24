import fs from 'node:fs/promises';
import sharp from 'sharp';
import { env } from '../config/env.js';

function escapeXml(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function buildTextOverlay({ categoryName, commercialCopy, phone, email }) {
  const title = escapeXml(categoryName.toUpperCase());
  const contact = escapeXml(`${phone}  |  ${email}`);
  const copyLines = String(commercialCopy || '')
    .match(/.{1,58}(\s|$)/g)
    ?.map((line) => escapeXml(line.trim()))
    .filter(Boolean) || [];
  const copyMarkup = copyLines
    .slice(0, 2)
    .map((line, index) => `<tspan x="72" dy="${index === 0 ? 0 : 34}">${line}</tspan>`)
    .join('');

  return Buffer.from(`
    <svg width="1080" height="1080" viewBox="0 0 1080 1080" xmlns="http://www.w3.org/2000/svg">
      <defs>
        <linearGradient id="fade" x1="0" x2="0" y1="0" y2="1">
          <stop offset="0%" stop-color="rgba(0,0,0,0.02)" />
          <stop offset="60%" stop-color="rgba(0,0,0,0.18)" />
          <stop offset="100%" stop-color="rgba(6,10,16,0.84)" />
        </linearGradient>
      </defs>
      <rect width="1080" height="1080" fill="url(#fade)" />
      <rect x="48" y="715" width="984" height="310" rx="28" fill="rgba(6,10,16,0.72)" />
      <text x="72" y="815" fill="#f8fafc" font-family="Arial, Helvetica, sans-serif" font-size="58" font-weight="700">${title}</text>
      <text x="72" y="874" fill="#d1d5db" font-family="Arial, Helvetica, sans-serif" font-size="28">${copyMarkup}</text>
      <text x="72" y="962" fill="#ffffff" font-family="Arial, Helvetica, sans-serif" font-size="30" font-weight="600">${contact}</text>
      <text x="72" y="1000" fill="#fca311" font-family="Arial, Helvetica, sans-serif" font-size="24" font-weight="700">NO LIMIT FINISH CARPENTRY</text>
    </svg>
  `);
}

async function loadOverlayBuffer(source, width) {
  if (!source) return null;

  const input = /^https?:\/\//i.test(source)
    ? Buffer.from(await (await fetch(source)).arrayBuffer())
    : await fs.readFile(source);

  return sharp(input).resize({ width }).png().toBuffer();
}

export async function generatePromotionalImage({ sourceImage, categoryName }) {
  const base = await sharp(sourceImage)
    .resize(1080, 1080, { fit: 'cover', position: 'centre' })
    .modulate({ brightness: 0.9, saturation: 1.05 })
    .png()
    .toBuffer();

  const overlays = [
    { input: buildTextOverlay({
        categoryName,
        commercialCopy: env.brandCopy,
        phone: env.brandPhone,
        email: env.brandEmail
      }), top: 0, left: 0 }
  ];

  try {
    const logo = await loadOverlayBuffer(env.logoPath, 280);
    if (logo) overlays.push({ input: logo, top: 44, left: 44 });
  } catch (_error) {
    // Keep generation working even if the logo asset is temporarily missing.
  }

  try {
    const badge = await loadOverlayBuffer(env.anniversaryBadgePath, 150);
    if (badge) overlays.push({ input: badge, top: 52, left: 874 });
  } catch (_error) {
    // Keep generation working even if the badge asset is temporarily missing.
  }

  return sharp(base).composite(overlays).png().toBuffer();
}
