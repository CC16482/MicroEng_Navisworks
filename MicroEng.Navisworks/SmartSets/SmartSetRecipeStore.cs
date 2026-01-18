using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;

namespace MicroEng.Navisworks.SmartSets
{
    public sealed class SmartSetRecipeStore
    {
        private readonly string _recipesDir;

        public SmartSetRecipeStore(string recipesDir)
        {
            if (string.IsNullOrWhiteSpace(recipesDir))
            {
                throw new ArgumentException("Invalid recipes directory.", nameof(recipesDir));
            }

            _recipesDir = recipesDir;
            Directory.CreateDirectory(_recipesDir);
        }

        public string RecipesDirectory => _recipesDir;

        public IEnumerable<string> ListRecipeFiles()
        {
            if (!Directory.Exists(_recipesDir))
            {
                yield break;
            }

            foreach (var file in Directory.EnumerateFiles(_recipesDir, "*.json").OrderBy(f => f))
            {
                yield return file;
            }
        }

        public SmartSetRecipe Load(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("Invalid recipe path.", nameof(path));
            }

            using (var fs = File.OpenRead(path))
            {
                var ser = new DataContractJsonSerializer(typeof(SmartSetRecipe));
                return (SmartSetRecipe)ser.ReadObject(fs);
            }
        }

        public void Save(SmartSetRecipe recipe, string path)
        {
            if (recipe == null) throw new ArgumentNullException(nameof(recipe));
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Invalid recipe path.", nameof(path));

            recipe.UpdatedUtc = DateTime.UtcNow;

            var tmp = path + ".tmp";
            using (var fs = File.Create(tmp))
            {
                var ser = new DataContractJsonSerializer(typeof(SmartSetRecipe));
                ser.WriteObject(fs, recipe);
            }

            if (File.Exists(path))
            {
                File.Delete(path);
            }

            File.Move(tmp, path);
        }

        public static string MakeSafeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                name = "Recipe";
            }

            foreach (var c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }

            return name.Trim();
        }

        public string GetDefaultPathForRecipe(string recipeName)
        {
            var safe = MakeSafeFileName(recipeName);
            return Path.Combine(_recipesDir, safe + ".json");
        }
    }
}
