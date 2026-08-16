using System;

namespace PresentationAssistant
{
    /// <summary>
    /// Counts consecutive invocations of the same command so the overlay can show
    /// "Scroll Line Down via Ctrl+Down Arrow ×9" instead of flashing nine times.
    /// The run is broken either by a different command or by
    /// <see cref="MultiplierTimeout"/> elapsing between two invocations.
    /// </summary>
    public class ShortcutDisplayStatistics
    {
        private DateTime _lastDisplayed;

        public ShortcutDisplayStatistics(int multiplierTimeoutInMS)
        {
            SetMultiplierTimeout(multiplierTimeoutInMS);
        }

        /// <summary>Length of the current run. Always at least 1 after <see cref="OnAction"/>.</summary>
        public int Multiplier { get; private set; } = 1;

        public string LastActionId { get; private set; }

        public TimeSpan MultiplierTimeout { get; private set; }

        public void SetMultiplierTimeout(int multiplierTimeoutInMS)
        {
            MultiplierTimeout = TimeSpan.FromMilliseconds(multiplierTimeoutInMS);
        }

        public void OnAction(string actionId)
        {
            var now = DateTime.UtcNow;

            if (actionId == LastActionId && now - _lastDisplayed < MultiplierTimeout)
            {
                ++Multiplier;
            }
            else
            {
                Multiplier = 1;
            }

            _lastDisplayed = now;
            LastActionId = actionId;
        }
    }
}
