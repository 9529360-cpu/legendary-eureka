/**
 * 用户偏好管理
 *
 * @module preferences
 */

import { UserPreferences, DEFAULT_PREFERENCES, PREFERENCES_STORAGE_KEY } from "../types/ui.types";

/**
 * 加载用户偏好
 */
export function loadUserPreferences(): UserPreferences {
  try {
    const stored = localStorage.getItem(PREFERENCES_STORAGE_KEY);
    if (stored) {
      return { ...DEFAULT_PREFERENCES, ...JSON.parse(stored) };
    }
  } catch (error) {
    console.warn("加载用户偏好失败:", error);
  }
  return DEFAULT_PREFERENCES;
}

/**
 * 保存用户偏好
 */
export function saveUserPreferences(prefs: UserPreferences): void {
  try {
    localStorage.setItem(PREFERENCES_STORAGE_KEY, JSON.stringify(prefs));
  } catch (error) {
    console.warn("保存用户偏好失败:", error);
  }
}
