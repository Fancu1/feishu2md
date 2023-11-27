package core

import (
	"context"
	"regexp"

	"github.com/Wsine/feishu2md/utils"
	"github.com/pkg/errors"
)

func GetDocsContent(url string, indentLevel int) (string, error) {
	configPath, err := GetConfigFilePath()
	utils.CheckErr(err)
	config, err := ReadConfigFromFile(configPath)
	utils.CheckErr(err)

	reg := regexp.MustCompile("^https://[a-zA-Z0-9-]+.(feishu.cn|larksuite.com)/(docx|wiki)/([a-zA-Z0-9]+)")
	matchResult := reg.FindStringSubmatch(url)
	if matchResult == nil || len(matchResult) != 4 {
		return "", errors.Errorf("Invalid feishu/larksuite URL format")
	}

	domain := matchResult[1]
	docType := matchResult[2]
	docToken := matchResult[3]

	ctx := context.WithValue(context.Background(), "output", config.Output)

	client := NewClient(
		config.Feishu.AppId, config.Feishu.AppSecret, domain,
	)

	// for a wiki page, we need to renew docType and docToken first
	if docType == "wiki" {
		node, err := client.GetWikiNodeInfo(ctx, docToken)
		utils.CheckErr(err)
		docType = node.ObjType
		docToken = node.ObjToken
	}

	docx, blocks, err := client.GetDocxContent(ctx, docToken)
	utils.CheckErr(err)

	parser := NewParser(ctx)

	markdown := parser.ParseDocxContent(docx, blocks, indentLevel)

	return markdown[1:], nil
}
