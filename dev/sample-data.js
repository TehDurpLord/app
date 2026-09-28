'use strict';
/**
 * Example parts for the local preview (npm run preview) and the screenshots.
 * Links point at supplier sites only so the preview looks realistic.
 */
const PARTS = [
  {
    name: 'M3 × 8 mm socket head cap screw', partNumber: '91292A112', category: 'Fasteners', location: 'Bin A-12',
    supplier: 'McMaster-Carr', quantity: 140, minQty: 50, reorderQty: 200, unit: 'ea', unitCost: 0.11,
    link: 'https://www.mcmaster.com/91292A112/',
  },
  {
    name: 'M3 hex nut, stainless', partNumber: '94150A325', category: 'Fasteners', location: 'Bin A-13',
    supplier: 'McMaster-Carr', quantity: 64, minQty: 50, reorderQty: 200, unit: 'ea', unitCost: 0.04,
    link: 'https://www.mcmaster.com/94150A325/',
  },
  {
    name: '6203-2RS sealed ball bearing', partNumber: '6203-2RS', category: 'Bearings', location: 'Shelf B-2',
    supplier: 'Grainger', quantity: 6, minQty: 4, reorderQty: 10, unit: 'ea', unitCost: 7.85,
    link: 'https://www.grainger.com/product/1ZGF8',
  },
  {
    name: 'V-belt A42', partNumber: 'A42', category: 'Belts', location: 'Rack C-1', supplier: 'Grainger',
    quantity: 3, minQty: 2, reorderQty: 4, unit: 'ea', unitCost: 12.4,
    link: 'https://www.grainger.com/product/3X538\nhttps://www.zoro.com/search?q=A42',
    notes: 'Conveyor 3 drive. Keep two spares on the rack.',
  },
  {
    name: 'Hydraulic filter element, 10 µm', partNumber: 'HF6553', category: 'Filters', location: 'Shelf D-4',
    supplier: 'Fleetguard', quantity: 3, minQty: 2, reorderQty: 6, unit: 'ea', unitCost: 38.9,
    link: 'https://www.fleetguard.com/', notes: 'Fits both presses on line 2.',
  },
  {
    name: 'Nitrile gloves, large', partNumber: 'S-9751L', category: 'Safety', location: 'Cabinet E',
    supplier: 'Uline', quantity: 5, minQty: 4, reorderQty: 10, unit: 'box', unitCost: 14.5,
    link: 'https://www.uline.com/Product/Detail/S-9751L/',
  },
  {
    name: '14 AWG THHN wire, black', partNumber: '14THHN-BK', category: 'Electrical', location: 'Reel rack',
    supplier: 'Home Depot', quantity: 320, minQty: 100, reorderQty: 500, unit: 'ft', unitCost: 0.16,
    link: 'https://www.homedepot.com/',
  },
  {
    name: '5 A fast-acting fuse, 5 × 20 mm', partNumber: '0217005.MXP', category: 'Electrical', location: 'Drawer F-3',
    supplier: 'Digi-Key', quantity: 14, minQty: 10, reorderQty: 50, unit: 'ea', unitCost: 0.38,
    link: 'https://www.digikey.com/',
  },
  {
    name: 'Cable ties, 8 in, black UV', partNumber: 'S-8746', category: 'Electrical', location: 'Drawer F-1',
    supplier: 'Uline', quantity: 850, minQty: 200, reorderQty: 1000, unit: 'ea', unitCost: 0.03,
    link: 'https://www.uline.com/Product/Detail/S-8746/',
  },
  {
    name: 'Proximity sensor, M12 PNP', partNumber: 'E2E-X4MD1', category: 'Sensors', location: 'Shelf B-5',
    supplier: 'AutomationDirect', quantity: 2, minQty: 1, reorderQty: 2, unit: 'ea', unitCost: 64,
    link: 'https://www.automationdirect.com/',
  },
  {
    name: 'Lithium grease, 14 oz cartridge', partNumber: 'LG-14', category: 'Lubricants', location: 'Cabinet E',
    supplier: 'Grainger', quantity: 9, minQty: 6, reorderQty: 12, unit: 'ea', unitCost: 6.2,
    link: 'https://www.grainger.com/',
  },
  {
    name: 'Shop towels, blue', category: 'Consumables', location: 'Cabinet E', supplier: 'Amazon',
    quantity: 11, minQty: 3, reorderQty: 6, unit: 'roll', link: 'https://www.amazon.com/',
  },
];

/**
 * Fills a freshly set-up app with parts and a few days of history, including
 * parts that are low, out of stock and on order.
 */
function seed(app) {
  const day = 24 * 3600 * 1000;
  const start = Date.now() - 3 * day;
  const at = (offset) => app.setTime(new Date(start + offset).toISOString());
  const byName = {};
  app.run('apiSaveSettings', {}, { recipients: 'purchasing@example.com, shop@example.com' });

  at(0);
  PARTS.forEach((part, i) => {
    at(i * 60 * 1000);
    byName[part.name] = app.run('apiSavePart', { actor: 'Jordan' }, part).part.id;
  });
  // Coworkers opening a shared link aren't identified by Google, so their typed name is logged.
  const use = (offset, actor, name, mode, amount, note) => {
    at(offset);
    const owner = app.env.active;
    app.env.active = '';
    app.run('apiAdjustStock', { actor }, { id: byName[name], mode, amount, note });
    app.env.active = owner;
  };
  use(1 * day + 3.2e6, 'Maria', 'M3 hex nut, stainless', 'remove', 26, 'Guard rebuild, line 1');
  use(1 * day + 9.1e6, 'Dev', 'V-belt A42', 'remove', 1, 'Conveyor 3');
  use(2 * day + 1.2e6, 'Maria', 'Nitrile gloves, large', 'remove', 3);
  at(2 * day + 1.5e6);
  app.run('apiSetOrdered', { actor: 'Jordan' }, byName['Nitrile gloves, large'], true);
  use(2 * day + 5.4e6, 'Dev', '5 A fast-acting fuse, 5 × 20 mm', 'remove', 7, 'Panel 4 repair');
  use(3 * day - 4.1e6, 'Maria', 'Cable ties, 8 in, black UV', 'remove', 100);
  use(3 * day - 2.6e6, 'Dev', 'V-belt A42', 'remove', 2, 'Conveyor 3');
  use(3 * day - 1.2e6, 'Jordan', '14 AWG THHN wire, black', 'add', 250);
  app.env.now = null;
  return byName;
}

module.exports = { PARTS, seed };
