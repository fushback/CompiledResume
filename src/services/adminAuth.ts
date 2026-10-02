/**
 * 6-Digit Admin PIN Authentication & LocalStorage Management
 * Ensures anyone can view the resume publicly, but any create, update,
 * or delete (CRUD) action requires entering a valid 6-digit PIN.
 */

const LOCAL_STORAGE_PIN_KEY = 'resume_admin_pin_6digit';

export function getStoredPin(): string | null {
  try {
    const pin = localStorage.getItem(LOCAL_STORAGE_PIN_KEY);
    return pin && pin.trim().length === 6 ? pin.trim() : null;
  } catch {
    return null;
  }
}

export function setStoredPin(pin: string): void {
  try {
    localStorage.setItem(LOCAL_STORAGE_PIN_KEY, pin.trim());
    window.dispatchEvent(new CustomEvent('resume_pin_changed', { detail: { pin: pin.trim() } }));
  } catch (err) {
    console.error('Failed to save PIN in localStorage:', err);
  }
}

export function clearStoredPin(): void {
  try {
    localStorage.removeItem(LOCAL_STORAGE_PIN_KEY);
    window.dispatchEvent(new CustomEvent('resume_pin_changed', { detail: { pin: null } }));
  } catch (err) {
    console.error('Failed to clear PIN from localStorage:', err);
  }
}

export function isCrudUnlocked(): boolean {
  return !!getStoredPin();
}

// Authorized PINs for browser verification
const CLIENT_VALID_PINS = new Set([
  '895193',
  '233235',
  '123456',
  '861808',
  '560100',
]);

/**
 * Validates a 6-digit PIN against the backend API with static CDN client fallback
 */
export async function verifyPin(pin: string): Promise<boolean> {
  const cleanPin = pin.trim();
  if (cleanPin.length !== 6 || !/^\d{6}$/.test(cleanPin)) {
    return false;
  }

  // 1. Try backend API if running full-stack Node server
  try {
    const res = await fetch('/api/auth/verify-pin', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ pin: cleanPin }),
    });

    const contentType = res.headers.get('content-type') || '';
    if (res.ok && contentType.includes('application/json')) {
      const data = await res.json();
      return !!data.valid;
    }
  } catch {
    // Fall through to client validation
  }

  // 2. Client-side validation for static Vercel hosting
  return CLIENT_VALID_PINS.has(cleanPin);
}
