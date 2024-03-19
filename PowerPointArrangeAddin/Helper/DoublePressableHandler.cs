using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PowerPointArrangeAddin.Helper {

    public class DoublePressableHandler {
        public DoublePressableHandler(bool enableDoublePress = true) {
            EnableDoublePress = enableDoublePress;
        }

        public bool EnableDoublePress { get; set; }

        private int _pressCheckRemaining;
        private bool _checkingFinished;
        private Point _cursorPosition;

        public void CheckPress(Action onPressed, Action onDoublePressed) {
            if (!EnableDoublePress) {
                onPressed?.Invoke();
                return;
            }

            if (_pressCheckRemaining > 0 && !_checkingFinished) {
                onDoublePressed?.Invoke();
                _pressCheckRemaining = 0;
                _checkingFinished = true;
                _cursorPosition = Point.Empty;
                return;
            }

            _pressCheckRemaining = Convert.ToInt32(Math.Ceiling(150.0 / 10)); // 150ms totally, each 10ms, check 15 times
            _checkingFinished = false;
            _cursorPosition = Cursor.Position;
            Task.Run(async () => {
                while (_pressCheckRemaining > 0 && _cursorPosition == Cursor.Position && !_checkingFinished) {
                    await Task.Delay(10);
                    _pressCheckRemaining--;
                }
                if ((_pressCheckRemaining <= 0 || _cursorPosition != Cursor.Position) && !_checkingFinished) {
                    onPressed?.Invoke();
                    _pressCheckRemaining = 0;
                    _checkingFinished = true;
                    _cursorPosition = Point.Empty;
                }
            });
        }
    }
}
