using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Settings;
using Newtonsoft.Json;
using ShowToolWindows.Model;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ShowToolWindows.Services
{
    /// <summary>
    /// Service for persisting tool window stashes.
    /// </summary>
    internal class StashSettingsService
    {
        private const string CollectionPath = "ShowToolWindows\\Stashes";
        private const string StashPrefix = "Stash_";

        private readonly WritableSettingsStore _settingsStore;

        /// <summary>
        /// Initializes a new instance of the <see cref="StashSettingsService"/> class.
        /// </summary>
        /// <param name="serviceProvider">The service provider used to access the Visual Studio settings store.</param>
        public StashSettingsService(IServiceProvider serviceProvider)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var settingsManager = new ShellSettingsManager(serviceProvider);
            _settingsStore = settingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);

            if (!_settingsStore.CollectionExists(CollectionPath))
            {
                _settingsStore.CreateCollection(CollectionPath);
            }
        }

        /// <summary>
        /// Saves all stashes to persistent settings, replacing any previously saved stashes.
        /// </summary>
        /// <param name="stashes">The stashes to persist.</param>
        public void SaveStashes(IEnumerable<ToolWindowStash> stashes)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var propertyNames = _settingsStore.GetPropertyNames(CollectionPath);
            foreach (var propertyName in propertyNames)
            {
                if (propertyName.StartsWith(StashPrefix))
                {
                    _settingsStore.DeleteProperty(CollectionPath, propertyName);
                }
            }

            int index = 0;
            foreach (var stash in stashes)
            {
                var key = StashPrefix + index++;
                var jsonValue = JsonConvert.SerializeObject(stash);
                _settingsStore.SetString(CollectionPath, key, jsonValue);
            }
        }

        /// <summary>
        /// Loads all previously saved stashes from persistent settings.
        /// </summary>
        /// <returns>A list of <see cref="ToolWindowStash"/> objects in the order they were saved.</returns>
        public IReadOnlyList<ToolWindowStash> LoadStashes()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var properties = _settingsStore.GetPropertyNames(CollectionPath);
            var stashProperties = properties
                .Where(p => p.StartsWith(StashPrefix))
                .OrderBy(p => int.TryParse(p.Substring(StashPrefix.Length), out int i) ? i : int.MaxValue)
                .ToList();

            var stashes = new List<ToolWindowStash>();
            foreach (var prop in stashProperties)
            {
                var jsonValue = _settingsStore.GetString(CollectionPath, prop);
                var stash = JsonConvert.DeserializeObject<ToolWindowStash>(jsonValue);
                stashes.Add(stash);
            }

            return stashes.AsReadOnly();
        }
    }
}