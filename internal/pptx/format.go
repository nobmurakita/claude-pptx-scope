package pptx

import (
	"fmt"
	"strings"
)

// formatAutoNum は自動番号を書式化する
func formatAutoNum(numType string, num int) string {
	switch numType {
	case "arabicPeriod":
		return fmt.Sprintf("%d.", num)
	case "arabicParenR":
		return fmt.Sprintf("%d)", num)
	case "alphaLcPeriod":
		return fmt.Sprintf("%s.", toLowerAlpha(num))
	case "alphaUcPeriod":
		return fmt.Sprintf("%s.", toUpperAlpha(num))
	case "romanLcPeriod":
		return fmt.Sprintf("%s.", toLowerRoman(num))
	case "romanUcPeriod":
		return fmt.Sprintf("%s.", toUpperRoman(num))
	default:
		return fmt.Sprintf("%d.", num)
	}
}

func toLowerAlpha(n int) string {
	if n < 1 {
		return fmt.Sprintf("%d", n)
	}
	var buf [8]byte
	i := len(buf)
	for n > 0 {
		n--
		i--
		buf[i] = byte('a' + n%26)
		n /= 26
	}
	return string(buf[i:])
}

func toUpperAlpha(n int) string {
	return strings.ToUpper(toLowerAlpha(n))
}

func toLowerRoman(n int) string {
	return strings.ToLower(toUpperRoman(n))
}

func toUpperRoman(n int) string {
	vals := []int{1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1}
	syms := []string{"M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"}
	var sb strings.Builder
	for i, v := range vals {
		for n >= v {
			sb.WriteString(syms[i])
			n -= v
		}
	}
	return sb.String()
}

// mapAlignment は OOXML 配置値を出力用文字列に変換する
func mapAlignment(algn string) string {
	switch algn {
	case "l":
		return "left"
	case "r":
		return "right"
	case "ctr":
		return "center"
	case "just":
		return "justify"
	default:
		return algn
	}
}

// mapVerticalAnchor は OOXML 垂直配置値を出力用文字列に変換する
func mapVerticalAnchor(anchor string) string {
	switch anchor {
	case "t":
		return "top"
	case "ctr":
		return "center"
	case "b":
		return "bottom"
	default:
		return ""
	}
}
