using Microsoft.VisualStudio.Settings;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Settings;
using Newtonsoft.Json;
using ShowToolWindows.Models;
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
        /// Saves all stashes to settings.
        /// </summary>
        public void SaveStashes(List<ToolWindowStash> stashes)
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

            for (int i = 0; i < stashes.Count; i++)
            {
                var stash = stashes[i];
                var key = StashPrefix + i;
                var jsonValue = JsonConvert.SerializeObject(stash);
                _settingsStore.SetString(CollectionPath, key, jsonValue);
            }
        }

        /// <summary>
        /// Loads all stashes from settings.
        /// </summary>
        public List<ToolWindowStash> LoadStashes()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            
            var properties = _settingsStore.GetPropertyNames(CollectionPath);
            var stashProperties = properties
                .Where(p => p.StartsWith(StashPrefix))
                .OrderBy(p => p)
                .ToList();

            var stashes = new List<ToolWindowStash>();
            foreach (var prop in stashProperties)
            {
                var jsonValue = _settingsStore.GetString(CollectionPath, prop);
                var stash = JsonConvert.DeserializeObject<ToolWindowStash>(jsonValue);
                stashes.Add(stash);
            }

            return stashes;
        }
    }
}
