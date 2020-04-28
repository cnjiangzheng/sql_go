package excel

type ExTools interface {
	ParseToSql() (string, error)
}
