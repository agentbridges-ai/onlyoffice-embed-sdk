import { ONLYOFFICE_LANG_KEY } from "../const";

export type OnlyOfficeLang =
  (typeof ONLYOFFICE_LANG_KEY)[keyof typeof ONLYOFFICE_LANG_KEY];

let currentLang: OnlyOfficeLang | null = null;

function detectOnlyOfficeLang(): OnlyOfficeLang {
  if (typeof navigator === "undefined") return ONLYOFFICE_LANG_KEY.EN;
  const requested = [navigator.language, ...(navigator.languages ?? [])]
    .find((language) => typeof language === "string" && language.length > 0)
    ?.toLowerCase();
  return requested?.startsWith("zh")
    ? ONLYOFFICE_LANG_KEY.ZH
    : ONLYOFFICE_LANG_KEY.EN;
}

export function getCurrentLang(): OnlyOfficeLang {
  return currentLang ?? detectOnlyOfficeLang();
}

export function setCurrentLang(lang: OnlyOfficeLang) {
  currentLang = lang;
}

export function getOnlyOfficeLang() {
  return getCurrentLang();
}
