import { useState } from 'react';

export type Region = 'EG' | 'KSA';

const REGION_KEY = 'cn-region';

/** BU codes (cfm_type) that belong to the Saudi Arabia region — everything else is Egypt. */
export const KSA_BU_CODES = ['AFW', 'AHJ', 'AKW', 'ALW', 'LCH', 'MKR', 'ADC', 'SNB'];

export function readStoredRegion(): Region | null {
  try {
    const v = localStorage.getItem(REGION_KEY);
    return v === 'EG' || v === 'KSA' ? v : null;
  } catch {
    return null;
  }
}

/** App-wide region choice (Egypt / Saudi Arabia), persisted to localStorage. Shared
 * between Layout (the selector modal + sidebar button) and any page that needs to
 * filter its data by region — lifted to App so both sides read the same value. */
export function useRegion(): [Region | null, (region: Region) => void] {
  const [region, setRegionState] = useState<Region | null>(readStoredRegion);

  function setRegion(next: Region) {
    try {
      localStorage.setItem(REGION_KEY, next);
    } catch {
      /* private browsing / storage disabled — region just won't persist */
    }
    setRegionState(next);
  }

  return [region, setRegion];
}
