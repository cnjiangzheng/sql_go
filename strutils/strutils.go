package strutils

import (
	"bytes"
	"fmt"
	"strings"
)

type StringBuilder struct {
	buf bytes.Buffer
}

func NewStringBuilder(s ...string) *StringBuilder {
	sb := StringBuilder{}
	sb.buf.WriteString(strings.Join(s, ""))
	return &sb
}

func (sb *StringBuilder) Len() int {
	return sb.buf.Len()
}

func (sb *StringBuilder) SetLen(n int) *StringBuilder {
	sb.buf.Truncate(n)
	return sb
}

func (sb *StringBuilder) Append(s string) *StringBuilder {
	sb.buf.WriteString(s)
	return sb
}

func (sb *StringBuilder) AppendFormat(f string, s ...interface{}) *StringBuilder {
	sb.buf.WriteString(fmt.Sprintf(f, s...))
	return sb
}

func (sb *StringBuilder) String() string {
	return sb.ToString()
}

func (sb *StringBuilder) ToString() string {
	return sb.buf.String()
}
