package core

import (
	"context"
	"fmt"
	"regexp"
	"strings"

	"github.com/Wsine/feishu2md/utils"
	"github.com/chyroc/lark"
	"github.com/olekukonko/tablewriter"
)

type Parser struct {
	ctx       context.Context
	ImgTokens []string
	blockMap  map[string]*lark.DocxBlock
	OnePage   bool
}

func NewParser(ctx context.Context, onePage bool) *Parser {
	return &Parser{
		ctx:       ctx,
		ImgTokens: make([]string, 0),
		blockMap:  make(map[string]*lark.DocxBlock),
		OnePage:   onePage,
	}
}

// =============================================================
// Parser utils
// =============================================================

var DocxCodeLang2MdStr = map[lark.DocxCodeLanguage]string{
	lark.DocxCodeLanguagePlainText:    "",
	lark.DocxCodeLanguageABAP:         "abap",
	lark.DocxCodeLanguageAda:          "ada",
	lark.DocxCodeLanguageApache:       "apache",
	lark.DocxCodeLanguageApex:         "apex",
	lark.DocxCodeLanguageAssembly:     "assembly",
	lark.DocxCodeLanguageBash:         "bash",
	lark.DocxCodeLanguageCSharp:       "csharp",
	lark.DocxCodeLanguageCPlusPlus:    "cpp",
	lark.DocxCodeLanguageC:            "c",
	lark.DocxCodeLanguageCOBOL:        "cobol",
	lark.DocxCodeLanguageCSS:          "css",
	lark.DocxCodeLanguageCoffeeScript: "coffeescript",
	lark.DocxCodeLanguageD:            "d",
	lark.DocxCodeLanguageDart:         "dart",
	lark.DocxCodeLanguageDelphi:       "delphi",
	lark.DocxCodeLanguageDjango:       "django",
	lark.DocxCodeLanguageDockerfile:   "dockerfile",
	lark.DocxCodeLanguageErlang:       "erlang",
	lark.DocxCodeLanguageFortran:      "fortran",
	lark.DocxCodeLanguageFoxPro:       "foxpro",
	lark.DocxCodeLanguageGo:           "go",
	lark.DocxCodeLanguageGroovy:       "groovy",
	lark.DocxCodeLanguageHTML:         "html",
	lark.DocxCodeLanguageHTMLBars:     "htmlbars",
	lark.DocxCodeLanguageHTTP:         "http",
	lark.DocxCodeLanguageHaskell:      "haskell",
	lark.DocxCodeLanguageJSON:         "json",
	lark.DocxCodeLanguageJava:         "java",
	lark.DocxCodeLanguageJavaScript:   "javascript",
	lark.DocxCodeLanguageJulia:        "julia",
	lark.DocxCodeLanguageKotlin:       "kotlin",
	lark.DocxCodeLanguageLateX:        "latex",
	lark.DocxCodeLanguageLisp:         "lisp",
	lark.DocxCodeLanguageLogo:         "logo",
	lark.DocxCodeLanguageLua:          "lua",
	lark.DocxCodeLanguageMATLAB:       "matlab",
	lark.DocxCodeLanguageMakefile:     "makefile",
	lark.DocxCodeLanguageMarkdown:     "markdown",
	lark.DocxCodeLanguageNginx:        "nginx",
	lark.DocxCodeLanguageObjective:    "objectivec",
	lark.DocxCodeLanguageOpenEdgeABL:  "openedge-abl",
	lark.DocxCodeLanguagePHP:          "php",
	lark.DocxCodeLanguagePerl:         "perl",
	lark.DocxCodeLanguagePostScript:   "postscript",
	lark.DocxCodeLanguagePower:        "powershell",
	lark.DocxCodeLanguageProlog:       "prolog",
	lark.DocxCodeLanguageProtoBuf:     "protobuf",
	lark.DocxCodeLanguagePython:       "python",
	lark.DocxCodeLanguageR:            "r",
	lark.DocxCodeLanguageRPG:          "rpg",
	lark.DocxCodeLanguageRuby:         "ruby",
	lark.DocxCodeLanguageRust:         "rust",
	lark.DocxCodeLanguageSAS:          "sas",
	lark.DocxCodeLanguageSCSS:         "scss",
	lark.DocxCodeLanguageSQL:          "sql",
	lark.DocxCodeLanguageScala:        "scala",
	lark.DocxCodeLanguageScheme:       "scheme",
	lark.DocxCodeLanguageScratch:      "scratch",
	lark.DocxCodeLanguageShell:        "shell",
	lark.DocxCodeLanguageSwift:        "swift",
	lark.DocxCodeLanguageThrift:       "thrift",
	lark.DocxCodeLanguageTypeScript:   "typescript",
	lark.DocxCodeLanguageVBScript:     "vbscript",
	lark.DocxCodeLanguageVisual:       "vbnet",
	lark.DocxCodeLanguageXML:          "xml",
	lark.DocxCodeLanguageYAML:         "yaml",
}

func renderMarkdownTable(data [][]string) string {
	builder := &strings.Builder{}
	table := tablewriter.NewWriter(builder)
	table.SetCenterSeparator("|")
	table.SetAutoWrapText(false)
	table.SetAutoFormatHeaders(false)
	table.SetAutoMergeCells(false)
	table.SetBorders(tablewriter.Border{Left: true, Top: false, Right: true, Bottom: false})
	table.SetHeader(data[0])
	table.AppendBulk(data[1:])
	table.Render()
	return builder.String()
}

// =============================================================
// Parse the new version of document (docx)
// =============================================================

func (p *Parser) ParseDocxContent(doc *lark.DocxDocument, blocks []*lark.DocxBlock) string {
	for _, block := range blocks {
		p.blockMap[block.BlockID] = block
	}

	entryBlock := p.blockMap[doc.DocumentID]
	return p.ParseDocxBlock(entryBlock, 0)
}

func (p *Parser) ParseDocxBlock(b *lark.DocxBlock, indentLevel int) string {
	buf := new(strings.Builder)
	buf.WriteString(strings.Repeat("\t", indentLevel))
	switch b.BlockType {
	case lark.DocxBlockTypePage:
		buf.WriteString(p.ParseDocxBlockPage(b))
	case lark.DocxBlockTypeText:
		buf.WriteString(p.ParseDocxBlockText(b.Text))
	case lark.DocxBlockTypeHeading1:
		buf.WriteString(p.processHeadingWithLink(b.Heading1, "#"))
	case lark.DocxBlockTypeHeading2:
		buf.WriteString(p.processHeadingWithLink(b.Heading2, "##"))
	case lark.DocxBlockTypeHeading3:
		buf.WriteString(p.processHeadingWithLink(b.Heading3, "###"))
	case lark.DocxBlockTypeHeading4:
		buf.WriteString(p.processHeadingWithLink(b.Heading4, "####"))
	case lark.DocxBlockTypeHeading5:
		buf.WriteString(p.processHeadingWithLink(b.Heading5, "#####"))
	case lark.DocxBlockTypeHeading6:
		buf.WriteString(p.processHeadingWithLink(b.Heading6, "######"))
	case lark.DocxBlockTypeHeading7:
		buf.WriteString(p.processHeadingWithLink(b.Heading7, "#######"))
	case lark.DocxBlockTypeHeading8:
		buf.WriteString(p.processHeadingWithLink(b.Heading8, "########"))
	case lark.DocxBlockTypeHeading9:
		buf.WriteString(p.processHeadingWithLink(b.Heading9, "#########"))
	case lark.DocxBlockTypeBullet:
		buf.WriteString(p.ParseDocxBlockBullet(b, indentLevel))
	case lark.DocxBlockTypeOrdered:
		buf.WriteString(p.ParseDocxBlockOrdered(b, indentLevel))
	case lark.DocxBlockTypeCode:
		buf.WriteString(p.ParseCodeBlock(b, indentLevel))
	case lark.DocxBlockTypeQuote:
		buf.WriteString("> ")
		buf.WriteString(p.ParseDocxBlockText(b.Quote))
	case lark.DocxBlockTypeEquation:
		buf.WriteString("$$\n")
		buf.WriteString(p.ParseDocxBlockText(b.Equation))
		buf.WriteString("\n$$\n")
	case lark.DocxBlockTypeTodo:
		if b.Todo.Style.Done {
			buf.WriteString("- [x] ")
		} else {
			buf.WriteString("- [ ] ")
		}
		buf.WriteString(p.ParseDocxBlockText(b.Todo))
	case lark.DocxBlockTypeDivider:
		buf.WriteString("---\n")
	case lark.DocxBlockTypeImage:
		buf.WriteString(p.ParseDocxBlockImage(b.Image))
	case lark.DocxBlockTypeTableCell:
		buf.WriteString(p.ParseDocxBlockTableCell(b))
	case lark.DocxBlockTypeTable:
		buf.WriteString(p.ParseDocxBlockTable(b.Table))
	case lark.DocxBlockTypeQuoteContainer:
		buf.WriteString(p.ParseDocxBlockQuoteContainer(b))
	default:
	}
	return buf.String()
}

func (p *Parser) processHeadingWithLink(bText *lark.DocxBlockText, headingPrefix string) string {
	headingContent := p.ParseDocxBlockText(bText)

	// if is a chapter link, load the chapter content
	if p.OnePage && p.IsChapterLink(headingContent) {
		chapterUrl := p.GetChapterLinkUrl(headingContent)
		chapterContent, err := GetDocsContent(chapterUrl)
		if err != nil {
			return ""
		}
		return headingPrefix + " " + chapterContent
	}
	return headingPrefix + " " + headingContent
}

func (p *Parser) ParseCodeBlock(b *lark.DocxBlock, indentLevel int) string {
	code := p.ParseDocxBlockText(b.Code)
	codeLanguage := DocxCodeLang2MdStr[b.Code.Style.Language]
	indentUnit := strings.Repeat("\t", indentLevel)

	buf := new(strings.Builder)
	buf.WriteString("\n")
	buf.WriteString(indentUnit + "```" + codeLanguage)
	buf.WriteString("\n")
	buf.WriteString(indentUnit + code)
	buf.WriteString(indentUnit + "```")
	buf.WriteString("\n")

	return buf.String()
}

func (p *Parser) ParseDocxBlockPage(b *lark.DocxBlock) string {
	buf := new(strings.Builder)

	buf.WriteString("# ")
	buf.WriteString(p.ParseDocxBlockText(b.Page))
	buf.WriteString("\n")

	for _, childId := range b.Children {
		childBlock := p.blockMap[childId]
		buf.WriteString(p.ParseDocxBlock(childBlock, 0))
		buf.WriteString("\n")
	}

	return buf.String()
}

func (p *Parser) ParseDocxBlockText(b *lark.DocxBlockText) string {
	buf := new(strings.Builder)
	numElem := len(b.Elements)
	for i := 0; i < numElem; i++ {
		inline := numElem > 1
		currentText := p.ParseDocxTextElement(b.Elements[i], inline)

		if strings.HasPrefix(currentText, "**") && strings.HasSuffix(currentText, "**") {
			trimmedText := strings.TrimPrefix(currentText, "**")
			trimmedText = strings.TrimSuffix(trimmedText, "**")
			trimmedText = strings.TrimSpace(trimmedText)
			currentText = "**" + trimmedText + "**"

			for i+1 < numElem {
				nextText := p.ParseDocxTextElement(b.Elements[i+1], inline)
				if strings.HasPrefix(nextText, "**") && strings.HasSuffix(nextText, "**") {
					nextTrimmed := strings.TrimPrefix(nextText, "**")
					nextTrimmed = strings.TrimSuffix(nextTrimmed, "**")
					nextTrimmed = strings.TrimSpace(nextTrimmed)

					currentText = strings.TrimSuffix(currentText, "**") + nextTrimmed + "**"
					i++
				} else {
					break
				}
			}
		}

		buf.WriteString(currentText)
	}

	buf.WriteString("\n")
	return buf.String()
}

// IsChapterLink checks if the string is a feishu other chapter link
// eg: [Chapter 1](https://xxx.feishu.cn/docx/xxx)
func (p *Parser) IsChapterLink(link string) bool {
	link = strings.TrimSpace(link)
	if strings.HasPrefix(link, "[") && strings.HasSuffix(link, ")") {
		closeBracketIndex := strings.Index(link, "]")
		openParenIndex := strings.Index(link, "(")

		if closeBracketIndex > 0 && openParenIndex == closeBracketIndex+1 {
			linkText := link[1:closeBracketIndex]
			url := link[openParenIndex+1 : len(link)-1]

			if linkText != "" && url != "" {
				return p.isValidUrl(url)
			}
		}
	}
	return false
}

// GetChapterLinkUrl returns the url of the chapter link
// eg: [Chapter 1](https://xxx.feishu.cn/docx/xxx)
func (p *Parser) GetChapterLinkUrl(link string) string {
	link = strings.TrimSpace(link)
	// split by "](", url is the second part
	linkParts := strings.Split(link, "](")
	if len(linkParts) != 2 {
		return ""
	}
	url := linkParts[1]
	// remove the last ")"
	url = strings.TrimSuffix(url, ")")
	return url
}

func (p *Parser) isValidUrl(url string) bool {
	// TODO: refactor
	reg := regexp.MustCompile("^https://[a-zA-Z0-9-]+.(feishu.cn|larksuite.com)/(docx|wiki)/([a-zA-Z0-9]+)")
	matchResult := reg.FindStringSubmatch(url)
	if matchResult == nil || len(matchResult) != 4 {
		return false
	}
	return true
}

func (p *Parser) ParseDocxTextElement(e *lark.DocxTextElement, inline bool) string {
	buf := new(strings.Builder)
	if e.TextRun != nil {
		buf.WriteString(p.ParseDocxTextElementTextRun(e.TextRun))
	}
	if e.MentionUser != nil {
		buf.WriteString(e.MentionUser.UserID)
	}
	if e.MentionDoc != nil {
		buf.WriteString(
			fmt.Sprintf("[%s](%s)", e.MentionDoc.Title, utils.UnescapeURL(e.MentionDoc.URL)))
	}
	if e.Equation != nil {
		symbol := "$$"
		if inline {
			symbol = "$"
		}
		buf.WriteString(symbol + strings.TrimSuffix(e.Equation.Content, "\n") + symbol)
	}
	return buf.String()
}

func (p *Parser) ParseDocxTextElementTextRun(tr *lark.DocxTextElementTextRun) string {
	buf := new(strings.Builder)
	postWrite := ""
	if style := tr.TextElementStyle; style != nil {
		useHTMLTags := NewConfig("", "").Output.UseHTMLTags
		if p.ctx.Value("output") != nil {
			useHTMLTags = p.ctx.Value("output").(OutputConfig).UseHTMLTags
		}
		if style.Bold {
			if useHTMLTags {
				buf.WriteString("<strong>")
				postWrite = "</strong>"
			} else {
				buf.WriteString("**")
				postWrite = "**"
			}
		} else if style.Italic {
			if useHTMLTags {
				buf.WriteString("<em>")
				postWrite = "</em>"
			} else {
				buf.WriteString("_")
				postWrite = "_"
			}
		} else if style.Strikethrough {
			if useHTMLTags {
				buf.WriteString("<del>")
				postWrite = "</del>"
			} else {
				buf.WriteString("~~")
				postWrite = "~~"
			}
		} else if style.Underline {
			buf.WriteString("<u>")
			postWrite = "</u>"
		} else if style.InlineCode {
			buf.WriteString("`")
			postWrite = "`"
		} else if link := style.Link; link != nil {
			buf.WriteString("[")
			postWrite = fmt.Sprintf("](%s)", utils.UnescapeURL(link.URL))
		}
	}
	buf.WriteString(tr.Content)
	buf.WriteString(postWrite)
	return buf.String()
}

func (p *Parser) ParseDocxBlockImage(img *lark.DocxBlockImage) string {
	buf := new(strings.Builder)
	buf.WriteString(fmt.Sprintf("![](%s)", img.Token))
	buf.WriteString("\n")
	p.ImgTokens = append(p.ImgTokens, img.Token)
	return buf.String()
}

func (p *Parser) ParseDocxWhatever(body *lark.DocBody) string {
	buf := new(strings.Builder)

	return buf.String()
}

func (p *Parser) ParseDocxBlockBullet(b *lark.DocxBlock, indentLevel int) string {
	buf := new(strings.Builder)

	buf.WriteString("- ")
	buf.WriteString(p.ParseDocxBlockText(b.Bullet))

	for _, childId := range b.Children {
		childBlock := p.blockMap[childId]
		buf.WriteString(p.ParseDocxBlock(childBlock, indentLevel+1))
	}

	return buf.String()
}

func (p *Parser) ParseDocxBlockOrdered(b *lark.DocxBlock, indentLevel int) string {
	buf := new(strings.Builder)

	// calculate order and indent level
	parent := p.blockMap[b.ParentID]
	order := 1
	for idx, child := range parent.Children {
		if child == b.BlockID {
			for i := idx - 1; i >= 0; i-- {
				if p.blockMap[parent.Children[i]].BlockType == lark.DocxBlockTypeOrdered {
					order += 1
				} else {
					break
				}
			}
			break
		}
	}

	buf.WriteString(fmt.Sprintf("%d. ", order))
	buf.WriteString(p.ParseDocxBlockText(b.Ordered))

	for _, childId := range b.Children {
		childBlock := p.blockMap[childId]
		buf.WriteString(p.ParseDocxBlock(childBlock, indentLevel+1))
	}

	return buf.String()
}

func (p *Parser) ParseDocxBlockTableCell(b *lark.DocxBlock) string {
	buf := new(strings.Builder)

	for _, child := range b.Children {
		block := p.blockMap[child]
		content := p.ParseDocxBlock(block, 0)
		buf.WriteString(content)
	}

	return buf.String()
}

func (p *Parser) ParseDocxBlockTable(t *lark.DocxBlockTable) string {
	// - First row as header
	// - Ignore cell merging
	var rows [][]string
	for i, blockId := range t.Cells {
		block := p.blockMap[blockId]
		cellContent := p.ParseDocxBlock(block, 0)
		cellContent = strings.ReplaceAll(cellContent, "\n", "")
		rowIndex := int64(i) / t.Property.ColumnSize
		if len(rows) < int(rowIndex)+1 {
			rows = append(rows, []string{})
		}
		rows[rowIndex] = append(rows[rowIndex], cellContent)
	}

	buf := new(strings.Builder)
	buf.WriteString(renderMarkdownTable(rows))
	buf.WriteString("\n")
	return buf.String()
}

func (p *Parser) ParseDocxBlockQuoteContainer(b *lark.DocxBlock) string {
	buf := new(strings.Builder)

	for _, child := range b.Children {
		block := p.blockMap[child]
		buf.WriteString("> ")
		buf.WriteString(p.ParseDocxBlock(block, 0))
	}

	return buf.String()
}
