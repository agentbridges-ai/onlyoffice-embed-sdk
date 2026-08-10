/**
 * The demo has twelve stable subframe slots. The host slugs are deliberately
 * ASCII so they can be used as DNS labels, while the Chinese names and emoji
 * make the slots easy to identify in the UI.
 */
export const CHINESE_ZODIAC_SLOTS = [
  { id: "rat", name: "鼠", emoji: "🐀" },
  { id: "ox", name: "牛", emoji: "🐂" },
  { id: "tiger", name: "虎", emoji: "🐅" },
  { id: "rabbit", name: "兔", emoji: "🐇" },
  { id: "dragon", name: "龙", emoji: "🐉" },
  { id: "snake", name: "蛇", emoji: "🐍" },
  { id: "horse", name: "马", emoji: "🐎" },
  { id: "goat", name: "羊", emoji: "🐐" },
  { id: "monkey", name: "猴", emoji: "🐒" },
  { id: "rooster", name: "鸡", emoji: "🐓" },
  { id: "dog", name: "狗", emoji: "🐕" },
  { id: "pig", name: "猪", emoji: "🐖" },
] as const;

export type ChineseZodiacSlot = (typeof CHINESE_ZODIAC_SLOTS)[number];

export const CHINESE_ZODIAC_SLOT_COUNT = CHINESE_ZODIAC_SLOTS.length;

export function getChineseZodiacSlot(index: number): ChineseZodiacSlot {
  const slot = CHINESE_ZODIAC_SLOTS[index];
  if (!slot) {
    throw new Error(`Chinese zodiac slot ${index + 1} does not exist`);
  }
  return slot;
}
