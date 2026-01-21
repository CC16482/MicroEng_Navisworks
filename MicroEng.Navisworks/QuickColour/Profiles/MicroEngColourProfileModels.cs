using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace MicroEng.Navisworks.QuickColour.Profiles
{
    public enum MicroEngColourProfileSource
    {
        QuickColour = 0,
        HierarchyBuilder = 1
    }

    public enum MicroEngColourApplyMode
    {
        Temporary = 0,
        Permanent = 1
    }

    public enum MicroEngPaletteKind
    {
        Deep = 0,
        Pastel = 1,
        CustomHue = 99
    }

    public static class MicroEngColourProfileSchema
    {
        public const int CurrentSchemaVersion = 1;
    }

    public sealed class MicroEngColourProfile
    {
        [JsonProperty("schemaVersion", Required = Required.Always)]
        public int SchemaVersion { get; set; } = MicroEngColourProfileSchema.CurrentSchemaVersion;

        [JsonProperty("profileId", Required = Required.Always)]
        public string Id { get; set; } = Guid.NewGuid().ToString("N");

        [JsonProperty("name", Required = Required.Always)]
        public string Name { get; set; } = "New Profile";

        [JsonProperty("source", Required = Required.Always)]
        public MicroEngColourProfileSource Source { get; set; } = MicroEngColourProfileSource.QuickColour;

        [JsonProperty("createdUtc", Required = Required.Always)]
        public DateTime CreatedUtc { get; set; } = DateTime.UtcNow;

        [JsonProperty("modifiedUtc", Required = Required.Always)]
        public DateTime ModifiedUtc { get; set; } = DateTime.UtcNow;

        [JsonProperty("scope", NullValueHandling = NullValueHandling.Include)]
        public MicroEngColourScope Scope { get; set; } = new MicroEngColourScope();

        [JsonProperty("generator", Required = Required.Always)]
        public MicroEngColourGenerator Generator { get; set; } = new MicroEngColourGenerator();

        [JsonProperty("outputs", NullValueHandling = NullValueHandling.Include)]
        public MicroEngColourOutputs Outputs { get; set; } = new MicroEngColourOutputs();

        [JsonProperty("rules", Required = Required.Always)]
        public List<MicroEngColourRule> Rules { get; set; } = new List<MicroEngColourRule>();
    }

    public sealed class MicroEngColourScope
    {
        [JsonProperty("kind")]
        public string Kind { get; set; } = "EntireModel";

        [JsonProperty("path")]
        public string Path { get; set; } = "";
    }

    public sealed class MicroEngColourGenerator
    {
        [JsonProperty("categoryName", Required = Required.Always)]
        public string CategoryName { get; set; } = "";

        [JsonProperty("propertyName", Required = Required.Always)]
        public string PropertyName { get; set; } = "";

        [JsonProperty("paletteName", Required = Required.Always)]
        public string PaletteName { get; set; } = "Deep";

        [JsonProperty("paletteKind", NullValueHandling = NullValueHandling.Include)]
        public MicroEngPaletteKind PaletteKind { get; set; } = MicroEngPaletteKind.Deep;

        [JsonProperty("customBaseColorHex", NullValueHandling = NullValueHandling.Include)]
        public string CustomBaseColorHex { get; set; } = "#FF6699";

        [JsonProperty("stableColors", Required = Required.Always)]
        public bool StableColors { get; set; } = true;

        [JsonProperty("seed", NullValueHandling = NullValueHandling.Include)]
        public string Seed { get; set; } = "";

        [JsonProperty("notes", NullValueHandling = NullValueHandling.Include)]
        public string Notes { get; set; } = "";
    }

    public sealed class MicroEngColourOutputs
    {
        [JsonProperty("createSearchSets")]
        public bool CreateSearchSets { get; set; }

        [JsonProperty("createSnapshots")]
        public bool CreateSnapshots { get; set; }

        [JsonProperty("folderPath")]
        public string FolderPath { get; set; } = "MicroEng/Quick Colour";

        [JsonProperty("setNamePrefix")]
        public string SetNamePrefix { get; set; } = "";
    }

    public sealed class MicroEngColourRule
    {
        [JsonProperty("enabled", Required = Required.Always)]
        public bool Enabled { get; set; } = true;

        [JsonProperty("value", Required = Required.Always)]
        public string Value { get; set; } = "";

        [JsonProperty("colorHex", Required = Required.Always)]
        public string ColorHex { get; set; } = "#FFFFFF";

        [JsonProperty("transparencyPercent", NullValueHandling = NullValueHandling.Ignore)]
        public double? TransparencyPercent { get; set; }

        [JsonProperty("count", NullValueHandling = NullValueHandling.Ignore)]
        public int? Count { get; set; }
    }
}
