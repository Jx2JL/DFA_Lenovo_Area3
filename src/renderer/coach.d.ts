interface CoachAPI {
  onShowTip: (callback: (message: string) => void) => void;
  onHideTip: (callback: () => void) => void;
}

interface Window {
  coach: CoachAPI;
}
