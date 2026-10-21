import unittest

from core.common import is_backup_sheet


class BackupSheetRuleTests(unittest.TestCase):
    def test_backup_sheet_names(self):
        for name in (
            # 名称含有效日期
            "20170119営業所用外部事業所設定",
            "20181213計画実績設定_v2.4",
            "帳票出力方法_20220929",
            "機能設計_20250315_bak",
            # bak / backup / old / copy（独立出现，不区分大小写）
            "利用者負担額調整_bak",
            "Sheet1-BAK",
            "bak2",
            "BACKUP_一覧",
            "一覧_old",
            "Copy of 一覧",
            # 日文 / 中文标记
            "機能設計（バックアップ）",
            "画面备份",
            "Sheet1のコピー",
            "一覧 副本",
            "旧画面定義",
            # Excel 复制工作表产生的后缀
            "一覧 (2)",
            "Sheet1 (12)",
        ):
            with self.subTest(name=name):
                self.assertTrue(is_backup_sheet(name))

    def test_regular_sheet_names_are_kept(self):
        for name in (
            "変更履歴",
            "利用者負担額調整",
            "Sheet1",
            "表紙",
            # 「新旧」「復旧」是正式资料里的常见用语
            "新旧対照表",
            "障害復旧手順",
            # 单词内部的 bak / old / copy 不算
            "Bakery",
            "Threshold設定",
            "Holder一覧",
            "Copyright",
            # 不带空格的「(2)」多为续页
            "機能一覧(2)",
            # 无效日期或更长的数字串
            "20251340設定",
            "画面ID_123456789",
        ):
            with self.subTest(name=name):
                self.assertFalse(is_backup_sheet(name))


if __name__ == "__main__":
    unittest.main()
