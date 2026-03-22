package search

import (
	"easyExcel/internal/config"
	"regexp"
	"strconv"
	"strings"
)

// Matcher matches cells against conditions
type Matcher struct {
	conditions []config.Condition
	logic      string // AND or OR
	compiledRegex map[int]*regexp.Regexp
}

// NewMatcher creates a new Matcher
func NewMatcher(conditions []config.Condition, logic string) (*Matcher, error) {
	m := &Matcher{
		conditions:    conditions,
		logic:         strings.ToUpper(logic),
		compiledRegex: make(map[int]*regexp.Regexp),
	}

	// Pre-compile regex patterns
	for i, cond := range conditions {
		if cond.Type == config.ConditionRegex {
			re, err := regexp.Compile(cond.Value)
			if err != nil {
				return nil, err
			}
			m.compiledRegex[i] = re
		}
	}

	return m, nil
}

// Match checks if a value matches the condition
func (m *Matcher) Match(value, conditionValue string, condType config.ConditionType) bool {
	value = strings.TrimSpace(value)
	conditionValue = strings.TrimSpace(conditionValue)

	switch condType {
	case config.ConditionEq:
		return value == conditionValue
	case config.ConditionContains:
		return strings.Contains(value, conditionValue)
	case config.ConditionStartsWith:
		return strings.HasPrefix(value, conditionValue)
	case config.ConditionEndsWith:
		return strings.HasSuffix(value, conditionValue)
	case config.ConditionGt:
		return compareNumeric(value, conditionValue) > 0
	case config.ConditionLt:
		return compareNumeric(value, conditionValue) < 0
	case config.ConditionGte:
		return compareNumeric(value, conditionValue) >= 0
	case config.ConditionLte:
		return compareNumeric(value, conditionValue) <= 0
	case config.ConditionRegex:
		return false // handled separately
	case config.ConditionEmpty:
		return value == ""
	default:
		return false
	}
}

// MatchWithRegex matches using a pre-compiled regex
func (m *Matcher) MatchWithRegex(value string, index int) bool {
	if re, ok := m.compiledRegex[index]; ok {
		return re.MatchString(value)
	}
	return false
}

// MatchCondition matches a single condition at the given index
func (m *Matcher) MatchCondition(value string, index int) bool {
	cond := m.conditions[index]
	if cond.Type == config.ConditionRegex {
		return m.MatchWithRegex(value, index)
	}
	return m.Match(value, cond.Value, cond.Type)
}

// Evaluate evaluates all conditions against the given values
// values is a map of column letter to cell value (e.g., "A" -> "value1")
func (m *Matcher) Evaluate(values map[string]string) bool {
	if len(m.conditions) == 0 {
		return false
	}

	result := false
	for i, cond := range m.conditions {
		cellValue := values[strings.ToUpper(cond.Column)]
		matches := m.MatchCondition(cellValue, i)

		if i == 0 {
			result = matches
		} else {
			// Determine logic with previous condition
			condLogic := strings.ToUpper(cond.Logic)
			if condLogic == "" {
				condLogic = m.logic
			}

			if condLogic == "OR" {
				result = result || matches
			} else { // AND
				result = result && matches
			}
		}
	}

	return result
}

// GetConditions returns the conditions
func (m *Matcher) GetConditions() []config.Condition {
	return m.conditions
}

// GetLogic returns the logic type
func (m *Matcher) GetLogic() string {
	return m.logic
}

// compareNumeric compares two numeric values
func compareNumeric(a, b string) int {
	// Try to parse as float64
	aFloat, errA := strconv.ParseFloat(a, 64)
	bFloat, errB := strconv.ParseFloat(b, 64)

	if errA != nil || errB != nil {
		// Fall back to string comparison
		if a > b {
			return 1
		} else if a < b {
			return -1
		}
		return 0
	}

	if aFloat > bFloat {
		return 1
	} else if aFloat < bFloat {
		return -1
	}
	return 0
}
