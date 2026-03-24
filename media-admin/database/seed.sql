insert into categories (name, slug)
values
  ('Trim', 'trim'),
  ('Kitchens', 'kitchens'),
  ('Decks', 'decks'),
  ('Wainscoting', 'wainscoting'),
  ('Stairs', 'stairs'),
  ('Ceiling', 'ceiling'),
  ('Pergola', 'pergola'),
  ('Fireplaces & Bars', 'fireplaces-bars'),
  ('Outside Doors & Windows', 'outside-doors-windows'),
  ('Port & Portal', 'port-portal'),
  ('Details', 'details')
on conflict (slug) do nothing;
