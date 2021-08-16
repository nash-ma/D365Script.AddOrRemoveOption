/// <reference path="../lib/@types/xrm/index.d.ts" />
namespace D365Script {
    export function AddOrRemoveOption(executionContext) {

        // 項目「案件ステータス」のオプション値：A
        const OPTION_VALUE_A = 100001001;
        // 項目「案件ステータス」のオプション値：B
        const OPTION_VALUE_B = 100002001;
        // 項目「案件ステータス」のオプション値：C
        const OPTION_VALUE_C = 100003001;
        // 項目「案件ステータス」のオプション値：D
        const OPTION_VALUE_D = 100004001;
        // 項目「案件ステータス」のオプション値：E
        const OPTION_VALUE_E = 100005001;

        // 項目「案件ステータス」のオプションラベル：A
        const OPTION_TEXT_A = "A";
        // 項目「案件ステータス」のオプションラベル：B
        const OPTION_TEXT_B = "B";
        // 項目「案件ステータス」のオプションラベル：C
        const OPTION_TEXT_C = "C";
        // 項目「案件ステータス」のオプションラベル：D
        const OPTION_TEXT_D = "D";
        // 項目「案件ステータス」のオプションラベル：E
        const OPTION_TEXT_E = "E";

        // 二つオプションラベル：はい
        const OPTION_TEXT_YES = "はい";
        // 二つオプションラベル：はい
        const OPTION_TEXT_NO = "いいえ";

        // 項目「案件ステータス」
        const COL_ANKEN_STATUS = "new_anken_status";
        // 項目「受注済」
        const COL_JUCHUSUMI = "new_juchusumi";

        // フォームコンテキストを取得
        var formContext = executionContext.getFormContext();

        // コントロール「案件ステータス」を取得
        var ctrlAnkenStatus = formContext.getControl(COL_ANKEN_STATUS);
        // コントロール「受注済」を取得
        var ctrlJuchusumi = formContext.getControl(COL_JUCHUSUMI);

        // 配列<整数>の宣言
        var valueAnkenStatus = new Array<number>();

        // 「受注済」が空白かをチェック
        if (isNullOrUndefined(ctrlJuchusumi)) {

            // 「案件ステータス」が空白か
            if (isNullOrUndefined(ctrlAnkenStatus)) {
                ctrlAnkenStatus.addOption({ text: OPTION_TEXT_C, value: OPTION_VALUE_C });
                ctrlAnkenStatus.addOption({ text: OPTION_TEXT_D, value: OPTION_VALUE_D });
                ctrlAnkenStatus.addOption({ text: OPTION_TEXT_E, value: OPTION_VALUE_E });
                // 「案件ステータス」の既定設定
                formContext.getAttribute(COL_ANKEN_STATUS).setValue(OPTION_VALUE_E);
            }
            return;
        }

        // 「受注済」の値を取得
        var valueJuchusumi = formContext.getControl(COL_JUCHUSUMI).getValue();

        // 「案件ステータス」が空白かをチェック
        if (!isNullOrUndefined(ctrlAnkenStatus)) {
            // 「案件ステータス」の値を取得
            valueAnkenStatus = formContext.getControl(COL_ANKEN_STATUS).getOptions().map(
                (x) =>
                    x.value);

        }

        // 「受注済」が”はい”か
        if (valueJuchusumi == OPTION_TEXT_YES) {

            // Aが存在するか
            if (valueAnkenStatus.indexOf(OPTION_VALUE_A) == -1)
                // Aを追加
                ctrlAnkenStatus.addOption({ text: OPTION_TEXT_A, value: OPTION_VALUE_A });

            // Bが存在するか
            if (valueAnkenStatus.indexOf(OPTION_VALUE_B) == -1)
                // Bを追加
                ctrlAnkenStatus.addOption({ text: OPTION_TEXT_B, value: OPTION_VALUE_B });

            // Cが存在するか
            if (valueAnkenStatus.indexOf(OPTION_VALUE_C) != -1)
                // Cを削除
                ctrlAnkenStatus.removeOption(OPTION_VALUE_C);

            // Dが存在するか
            if (valueAnkenStatus.indexOf(OPTION_VALUE_D) != -1)
                // Dを削除
                ctrlAnkenStatus.removeOption(OPTION_VALUE_D);

            // Eが存在するか
            if (valueAnkenStatus.indexOf(OPTION_VALUE_E) != -1)
                // Eを削除
                ctrlAnkenStatus.removeOption(OPTION_VALUE_E);

            // CDEか
            if (formContext.getAttribute(COL_ANKEN_STATUS).getValue() > OPTION_VALUE_B)
                // 「案件ステータス」の既定設定
                formContext.getAttribute(COL_ANKEN_STATUS).setValue(OPTION_VALUE_B);

        } else {

            // Aが存在するか
            if (valueAnkenStatus.indexOf(OPTION_VALUE_A) != -1)
                // Aを削除
                ctrlAnkenStatus.removeOption(OPTION_VALUE_A);

            // Bが存在するか
            if (valueAnkenStatus.indexOf(OPTION_VALUE_B) != -1)
                // Bを削除
                ctrlAnkenStatus.removeOption(OPTION_VALUE_B);

            // Cが存在するか
            if (valueAnkenStatus.indexOf(OPTION_VALUE_C) == -1)
                // Cを追加
                ctrlAnkenStatus.addOption({ text: OPTION_TEXT_C, value: OPTION_VALUE_C });

            // Dが存在するか
            if (valueAnkenStatus.indexOf(OPTION_VALUE_D) == -1)
                // Dを追加
                ctrlAnkenStatus.addOption({ text: OPTION_TEXT_D, value: OPTION_VALUE_D });

            // Eが存在するか
            if (valueAnkenStatus.indexOf(OPTION_VALUE_E) == -1)
                // Eを追加
                ctrlAnkenStatus.addOption({ text: OPTION_TEXT_E, value: OPTION_VALUE_E });

            // ABか
            if (formContext.getAttribute(COL_ANKEN_STATUS).getValue() < OPTION_VALUE_C)
                // 「案件ステータス」の既定設定
                formContext.getAttribute(COL_ANKEN_STATUS).setValue(OPTION_VALUE_E);

        }

    }

    // 空白のチェック
    function isNullOrUndefined<T>(object: T | null | undefined): object is T {
        return <T>object == null || <T>object == undefined;
    }

}