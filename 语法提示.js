(function () {
  Office.initialize = function () {
    // 在 Office 初始化时执行此任务。
    Word.run(function (context) {
      // 在此处编写您的代码。

      // 注册内容更改事件，以在文档内容更改时进行检查。
      var body = context.document.body;
      body.onTextChanged.add(function () {
        // 获取所有拼写错误并判断其语法是否有误
        var spellingErrors = context.document.body.getSpellingErrorRanges();
        for (var i = 0; i < spellingErrors.items.length; i++) {
          var errorRange = spellingErrors.items[i];

          if (errorRange.text.indexOf("’") !== -1) {
            // 不要将缩写识别为语法错误，例如“you’re”和“they’re”
            continue;
          }

          // 检查语法错误
          var grammarErrors = errorRange.getGrammarErrorRanges();
          for (var j = 0; j < grammarErrors.items.length; j++) {
            var grammarError = grammarErrors.items[j];
            var errorType = grammarError.getGrammarErrorType();

            // 根据错误类型生成适当的消息
            var message;
            switch (errorType) {
              case "accord":
                message = "主谓不一致";
                break;
              case "hyphens":
                message = "连字符用法错误";
                break;
              case "modal-verb":
                message = "情态动词用法错误";
                break;
              // 添加其他错误类型...
              default:
                message = "语法错误";
                break;
            }

            // 将消息添加到错误位置
            grammarError.insertText(message, "Replace");
          }
        }

        return context.sync();
      });

      return context.sync();
    });
  };
})();
