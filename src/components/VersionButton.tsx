import { useState } from "react";
import { CURRENT_VERSION, VERSION_HISTORY } from "../lib/version";
import { IconX } from "./Icons";

export function VersionButton() {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <>
      {/* Кнопка версии */}
      <button
        type="button"
        onClick={() => setIsOpen(true)}
        className="mt-8 inline-flex items-center gap-2 rounded-[6px] border border-line/40 bg-paper/50 px-3 py-1.5 font-mono text-[11px] text-ink-soft/60 transition-all hover:border-green/40 hover:bg-white hover:text-green-deep"
        title="История версий"
      >
        <span>Версия {CURRENT_VERSION.version}</span>
        <span className="text-ink-soft/40">от {CURRENT_VERSION.date}</span>
      </button>

      {/* Модальное окно с историей версий */}
      {isOpen && (
        <div
          className="fixed inset-0 z-50 grid place-items-center bg-ink/40 p-4"
          onClick={() => setIsOpen(false)}
        >
          <div
            className="card max-h-[80vh] w-full max-w-2xl overflow-hidden"
            onClick={(e) => e.stopPropagation()}
          >
            <header className="flex items-center gap-3 border-b border-line bg-green-mist px-5 py-4">
              <div className="min-w-0 flex-1">
                <p className="label-caps">история изменений</p>
                <h2 className="mt-1 font-display text-[18px] font-semibold tracking-tight text-ink">
                  Стабильные версии
                </h2>
              </div>
              <button
                type="button"
                onClick={() => setIsOpen(false)}
                className="grid h-8 w-8 place-items-center rounded-[6px] border border-line text-ink-soft transition-colors hover:border-red hover:text-red"
              >
                <IconX size={16} />
              </button>
            </header>

            <div className="max-h-[60vh] overflow-y-auto scroll-slim px-5 py-4">
              {VERSION_HISTORY.map((v, idx) => (
                <div key={v.version} className={idx > 0 ? "mt-6 border-t border-line pt-6" : ""}>
                  <div className="flex items-baseline gap-3">
                    <span className="font-display text-[16px] font-bold text-green-deep">
                      v{v.version}
                    </span>
                    <span className="font-mono text-[12px] text-ink-soft">{v.date}</span>
                    {idx === 0 && (
                      <span className="rounded-[4px] bg-green px-2 py-0.5 font-mono text-[10px] font-semibold text-white">
                        ТЕКУЩАЯ
                      </span>
                    )}
                  </div>
                  <ul className="mt-3 space-y-1.5">
                    {v.changes.map((change, i) => (
                      <li key={i} className="flex gap-2 text-[13px] leading-relaxed text-ink-soft">
                        <span className="mt-0.5 text-green">◆</span>
                        <span>{change}</span>
                      </li>
                    ))}
                  </ul>
                </div>
              ))}
            </div>

            <footer className="border-t border-line bg-paper/70 px-5 py-3 text-center font-mono text-[11px] text-ink-soft">
              Всего версий: {VERSION_HISTORY.length}
            </footer>
          </div>
        </div>
      )}
    </>
  );
}
