package config

import (
	"encoding/json"
	"fmt"
	"os"
	"regexp"
	"strconv"
	"strings"
)

// ConditionType represents the type of search condition
type ConditionType string

const (
	ConditionEq        ConditionType = "eq"
	ConditionContains  ConditionType = "contains"
	ConditionStartsWith ConditionType = "startswith"
	ConditionEndsWith  ConditionType = "endswith"
	ConditionGt        ConditionType = "gt"
	ConditionLt        ConditionType = "lt"
	ConditionGte       ConditionType = "gte"
	ConditionLte       ConditionType = "lte"
	ConditionRegex     ConditionType = "regex"
	ConditionEmpty     ConditionType = "empty"
)

// Condition represents a single search condition
type Condition struct {
	Column string        `json:"column"`
	Type   ConditionType `json:"type"`
	Value  string       `json:"value"`
	Logic  string       `json:"logic,omitempty"` // AND or OR, for combining with previous condition
}

// HighlightConfig represents highlight settings
type HighlightConfig struct {
	Cell string `json:"cell"`
	Row  string `json:"row"`
}

// Config represents the query configuration
type Config struct {
	Conditions []Condition    `json:"conditions"`
	Logic      string         `json:"logic"` // AND or OR
	Highlight  HighlightConfig `json:"highlight"`
}

// DefaultConfig returns a default configuration
func DefaultConfig() *Config {
	return &Config{
		Logic: "AND",
		Highlight: HighlightConfig{
			Cell: "#FFFF00",
			Row:  "#FF9900",
		},
	}
}

// LoadConfig loads configuration from a JSON file
func LoadConfig(path string) (*Config, error) {
	data, err := os.ReadFile(path)
	if err != nil {
		return nil, fmt.Errorf("failed to read config file: %w", err)
	}

	var cfg Config
	if err := json.Unmarshal(data, &cfg); err != nil {
		return nil, fmt.Errorf("failed to parse config file: %w", err)
	}

	return &cfg, nil
}

// SaveConfig saves configuration to a JSON file
func SaveConfig(cfg *Config, path string) error {
	data, err := json.MarshalIndent(cfg, "", "  ")
	if err != nil {
		return fmt.Errorf("failed to marshal config: %w", err)
	}
	if err := os.WriteFile(path, data, 0644); err != nil {
		return fmt.Errorf("failed to write config file: %w", err)
	}
	return nil
}

// Validate validates the configuration
func (c *Config) Validate() error {
	if len(c.Conditions) == 0 {
		return fmt.Errorf("no conditions specified")
	}

	logic := strings.ToUpper(c.Logic)
	if logic != "AND" && logic != "OR" {
		return fmt.Errorf("invalid logic: %s, must be AND or OR", c.Logic)
	}

	for i, cond := range c.Conditions {
		if cond.Column == "" {
			return fmt.Errorf("condition %d: column is required", i+1)
		}
		if err := validateConditionType(cond.Type); err != nil {
			return fmt.Errorf("condition %d: %w", i+1, err)
		}
		if cond.Type != ConditionEmpty && cond.Value == "" {
			return fmt.Errorf("condition %d: value is required for type %s", i+1, cond.Type)
		}
	}

	return nil
}

func validateConditionType(t ConditionType) error {
	validTypes := []ConditionType{
		ConditionEq, ConditionContains, ConditionStartsWith,
		ConditionEndsWith, ConditionGt, ConditionLt,
		ConditionGte, ConditionLte, ConditionRegex, ConditionEmpty,
	}

	for _, valid := range validTypes {
		if t == valid {
			return nil
		}
	}
	return fmt.Errorf("invalid condition type: %s", t)
}

// ParseColor parses a color string to RGB values
func ParseColor(color string) (r, g, b int, err error) {
	color = strings.TrimPrefix(color, "#")
	if len(color) == 3 {
		color = string(color[0]) + string(color[0]) +
			string(color[1]) + string(color[1]) +
			string(color[2]) + string(color[2])
	}
	if len(color) != 6 {
		return 0, 0, 0, fmt.Errorf("invalid color format: %s", color)
	}

	var rgb uint64
	rgb, err = strconv.ParseUint(color, 16, 32)
	if err != nil {
		return 0, 0, 0, fmt.Errorf("invalid color format: %s", color)
	}

	r = int((rgb >> 16) & 0xFF)
	g = int((rgb >> 8) & 0xFF)
	b = int(rgb & 0xFF)
	return r, g, b, nil
}

// CompileRegex compiles a regex pattern and returns it
func CompileRegex(pattern string) (*regexp.Regexp, error) {
	return regexp.Compile(pattern)
}

// ColorNameToHex converts a color name to hex
func ColorNameToHex(name string) string {
	colors := map[string]string{
		"red":    "#FF0000",
		"green":  "#00FF00",
		"blue":   "#0000FF",
		"yellow": "#FFFF00",
		"orange": "#FF9900",
		"purple": "#9900FF",
		"pink":   "#FF66FF",
		"gray":   "#808080",
		"white":  "#FFFFFF",
		"black":  "#000000",
	}

	if hex, ok := colors[strings.ToLower(name)]; ok {
		return hex
	}
	return name
}
