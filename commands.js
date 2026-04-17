// 初期化
Office.onReady(function () {
    // UIを持たないコマンドのため、初期化時は特に処理なし
});

// 送信ボタンが押されたときに実行される関数
function checkBodyOnSend(event) {
    // 非同期でメール本文（テキスト）を取得
    Office.context.mailbox.item.body.getAsync(
        "text",
        { asyncContext: event },
        function (asyncResult) {
            var event = asyncResult.asyncContext;

            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                var bodyText = asyncResult.value.toLowerCase();

                // 本文に "box.com" が含まれているか判定
                if (bodyText.indexOf("box.com") !== -1) {
                    // 含まれている場合は送信をブロック
                    event.completed({ 
                        allowEvent: false, 
                        errorMessage: "【セキュリティ警告】Boxのリンクが含まれています。アクセス権限を確認するか、リンクを削除してから再送信してください。" 
                    });
                } else {
                    // 問題なければ送信許可
                    event.completed({ allowEvent: true });
                }
            } else {
                // エラーで本文が取得できなかった場合は、業務影響を避けるため送信を許可
                event.completed({ allowEvent: true });
            }
        }
    );
}

// Outlookがこの関数を認識できるように関連付け（必須）
Office.actions.associate("checkBodyOnSend", checkBodyOnSend);