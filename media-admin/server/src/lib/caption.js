const CATEGORY_PHRASES = {
  trim: 'Precision trim details that instantly elevate the room.',
  kitchens: 'Custom kitchen craftsmanship built for daily life and a premium finish.',
  decks: 'Outdoor carpentry designed for comfort, durability, and curb appeal.',
  wainscoting: 'Wall paneling that adds depth, texture, and timeless character.',
  stairs: 'Stair finishes and details completed with clean alignment and strong visual impact.',
  ceiling: 'Ceiling details that transform plain space into standout design.',
  pergola: 'Custom pergola work that extends the home with style and shade.',
  'fireplaces-bars': 'Signature fireplace and bar details that create a focal point worth sharing.',
  'outside-doors-windows': 'Exterior trim and openings installed with lasting curb appeal in mind.',
  'port-portal': 'Entry details that make the first impression count.',
  details: 'Finishing details that show the difference between standard work and standout work.'
};

export function generateCaption({ categoryName, categorySlug, phone, email }) {
  const categoryLine =
    CATEGORY_PHRASES[categorySlug] ||
    `${categoryName} work delivered with craftsmanship, clean execution, and attention to detail.`;

  return [
    `${categoryName} by No Limit.`,
    categoryLine,
    `Ready to upgrade your space? Call ${phone} or email ${email}.`
  ].join(' ');
}
