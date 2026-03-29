package pptx

import (
	"testing"
)

func TestCleanPath(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"ppt/slides/../media/image1.png", "ppt/media/image1.png"},
		{"ppt/slides/slide1.xml", "ppt/slides/slide1.xml"},
		{"ppt/./slides/slide1.xml", "ppt/slides/slide1.xml"},
		{"a/b/c/../../d", "a/d"},
		{"", ""},
	}
	for _, tt := range tests {
		got := cleanPath(tt.input)
		if got != tt.want {
			t.Errorf("cleanPath(%q) = %q, want %q", tt.input, got, tt.want)
		}
	}
}

func TestResolveRelTarget(t *testing.T) {
	tests := []struct {
		basePath string
		target   string
		want     string
	}{
		{"ppt", "slides/slide1.xml", "ppt/slides/slide1.xml"},
		{"ppt", "/ppt/slides/slide1.xml", "ppt/slides/slide1.xml"},
		{"ppt/slides", "../media/image1.png", "ppt/media/image1.png"},
	}
	for _, tt := range tests {
		got := resolveRelTarget(tt.basePath, tt.target)
		if got != tt.want {
			t.Errorf("resolveRelTarget(%q, %q) = %q, want %q", tt.basePath, tt.target, got, tt.want)
		}
	}
}
