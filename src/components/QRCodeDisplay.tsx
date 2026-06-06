/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useEffect, useState } from 'react';
import QRCode from 'qrcode';
import { QrCode, Copy, Check } from 'lucide-react';

interface QRCodeDisplayProps {
  value: string;
  size?: number;
  darkColor?: string; // Hex for darker modules
  lightColor?: string; // Hex for background
  showCopyButton?: boolean;
}

export const QRCodeDisplay: React.FC<QRCodeDisplayProps> = ({
  value,
  size = 180,
  darkColor = '#0f172a', // deep slate slate-900
  lightColor = '#ffffff', // white
  showCopyButton = true
}) => {
  const [qrSrc, setQrSrc] = useState<string>('');
  const [copied, setCopied] = useState<boolean>(false);

  useEffect(() => {
    if (!value) return;

    QRCode.toDataURL(
      value,
      {
        width: size,
        margin: 1,
        color: {
          dark: darkColor,
          light: lightColor
        }
      },
      (err, url) => {
        if (err) {
          console.error('Failed to generate QR Code:', err);
          return;
        }
        setQrSrc(url);
      }
    );
  }, [value, size, darkColor, lightColor]);

  const handleCopyLink = async () => {
    try {
      await navigator.clipboard.writeText(value);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch (err) {
      console.warn('Failed to copy to clipboard', err);
    }
  };

  return (
    <div className="flex flex-col items-center justify-center p-2 rounded-xl bg-white border border-slate-200 shadow-sm transition-all hover:shadow-md">
      {qrSrc ? (
        <img
          src={qrSrc}
          alt={`QR Code para ${value}`}
          style={{ width: size, height: size }}
          className="rounded-lg select-none"
          referrerPolicy="no-referrer"
        />
      ) : (
        <div
          style={{ width: size, height: size }}
          className="flex flex-col items-center justify-center text-slate-400 bg-slate-50 rounded-lg animate-pulse"
        >
          <QrCode className="w-8 h-8 opacity-50 mb-1" />
          <span className="text-xs">Gerando QR...</span>
        </div>
      )}
      
      {showCopyButton && (
        <button
          onClick={handleCopyLink}
          className="mt-2 text-xs flex items-center justify-center gap-1 text-slate-600 hover:text-slate-900 active:scale-95 transition-all bg-slate-150 px-3 py-1.5 rounded-lg w-full font-medium"
        >
          {copied ? (
            <>
              <Check className="w-3.5 h-3.5 text-emerald-600" />
              <span className="text-emerald-700">Link Copiado!</span>
            </>
          ) : (
            <>
              <Copy className="w-3.5 h-3.5 text-slate-500" />
              <span>Copiar Link de Acesso</span>
            </>
          )}
        </button>
      )}
    </div>
  );
};
