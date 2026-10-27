import { KeyboardEvent, PointerEvent, useEffect, useRef, useState } from "react";

const desktopPreference = "monoreport.playground.splitter.desktop";
const mobilePreference = "monoreport.playground.splitter.mobile";

function readPreference(mobile: boolean) {
  const value = Number(localStorage.getItem(mobile ? mobilePreference : desktopPreference));
  if (!Number.isFinite(value) || value <= 0) return mobile ? 50 : 45;
  return Math.max(mobile ? 20 : 25, Math.min(mobile ? 80 : 75, value));
}

export function Splitter() {
  const [mobile, setMobile] = useState(() => window.matchMedia("(max-width: 850px)").matches);
  const [value, setValue] = useState(() => readPreference(mobile));
  const [range, setRange] = useState(() => ({ min: mobile ? 20 : 25, max: mobile ? 80 : 75 }));
  const dragging = useRef(false);

  useEffect(() => {
    const media = window.matchMedia("(max-width: 850px)");
    const restore = (nextMobile: boolean) => {
      setMobile(nextMobile);
      apply(readPreference(nextMobile), nextMobile, false);
    };
    const update = () => {
      restore(media.matches);
    };
    update();
    media.addEventListener("change", update);
    window.addEventListener("resize", update);
    return () => {
      media.removeEventListener("change", update);
      window.removeEventListener("resize", update);
      document.documentElement.style.removeProperty("--source-height");
      document.documentElement.style.removeProperty("--left-width");
    };
  }, []); // Layout preferences are read again whenever the viewport changes.

  function apply(next: number, isMobile = mobile, persist = true) {
    const bounds = document.querySelector(".layout")?.getBoundingClientRect();
    const size = bounds ? (isMobile ? bounds.height : bounds.width) : 0;
    const gap = isMobile ? 8 : 7;
    const minimum = isMobile ? 160 : 320;
    const otherMinimum = isMobile ? 160 : 360;
    const percentageMinimum = isMobile ? 20 : 25;
    const percentageMaximum = isMobile ? 80 : 75;
    const min = Math.max(percentageMinimum, size > 0 ? minimum / size * 100 : percentageMinimum);
    const max = Math.min(percentageMaximum, size > 0 ? (size - gap - otherMinimum) / size * 100 : percentageMaximum);
    const clamped = Math.max(min, Math.min(max, next));
    setRange({ min, max });
    setValue(clamped);
    document.documentElement.style.setProperty(isMobile ? "--source-height" : "--left-width", `${clamped}%`);
    if (persist) localStorage.setItem(isMobile ? mobilePreference : desktopPreference, String(clamped));
  }

  function pointerMove(event: PointerEvent<HTMLDivElement>) {
    if (!dragging.current) return;
    const layout = event.currentTarget.parentElement;
    if (!layout) return;
    const bounds = layout.getBoundingClientRect();
    apply(mobile ? ((event.clientY - bounds.top) / bounds.height) * 100 : ((event.clientX - bounds.left) / bounds.width) * 100);
  }

  function keyDown(event: KeyboardEvent<HTMLDivElement>) {
    const delta = mobile ? (event.key === "ArrowDown" ? 3 : event.key === "ArrowUp" ? -3 : 0) : (event.key === "ArrowRight" ? 3 : event.key === "ArrowLeft" ? -3 : 0);
    if (!delta) return;
    event.preventDefault();
    apply(value + delta);
  }

  return <div
    className={`splitter ${mobile ? "horizontal" : "vertical"}`}
    role="separator"
    aria-label="Resize input and preview"
    aria-orientation={mobile ? "horizontal" : "vertical"}
    aria-valuemin={Math.ceil(range.min)}
    aria-valuemax={Math.floor(range.max)}
    aria-valuenow={Math.round(value)}
    tabIndex={0}
    onKeyDown={keyDown}
    onPointerDown={event => { dragging.current = true; event.currentTarget.setPointerCapture(event.pointerId); }}
    onPointerMove={pointerMove}
    onPointerUp={() => { dragging.current = false; }}
    onPointerCancel={() => { dragging.current = false; }}
  />;
}
