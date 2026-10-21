import unittest

from core.common import find_backup_sheets, find_clean_targets, find_contained_sheets, has_backup_marker


class BackupMarkerTests(unittest.TestCase):
    def test_names_with_backup_markers(self):
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
            # 不带空格、全角括号，或后面再接日期的复制后缀
            "機能一覧(2)",
            "一覧（３）",
            "個別支援計画書_様式１（基本情報）(2)_20221005",
            # 带分隔符或括号的日期
            "概要(2018-02-26)",
            "概要_2018.02.26",
        ):
            with self.subTest(name=name):
                self.assertTrue(has_backup_marker(name))

    def test_names_without_backup_markers(self):
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
            # 无效日期或更长的数字串
            "20251340設定",
            "画面ID_123456789",
        ):
            with self.subTest(name=name):
                self.assertFalse(has_backup_marker(name))


class FindBackupSheetsTests(unittest.TestCase):
    def test_dated_sheets_without_original_are_kept(self):
        # SSPLG1100_個別支援計画書情報.xlsx：没有 基本情報 等原表，带日期的就是在用的表
        self.assertEqual(find_backup_sheets([
            "変更履歴", "個別支援計画書情報", "基本情報_20241003", "具体的な支援目標及び支援計画_20241003",
        ]), [])
        # 只有一个日期版本时它就是在用的表
        self.assertEqual(find_backup_sheets(["表紙", "基本情報_20241003"]), [])

    def test_older_dated_versions_without_original_are_backups(self):
        # 多个日期版本都没有原表：保留最新的一个，更早的视为备份
        self.assertEqual(find_backup_sheets(["基本情報_20241003", "基本情報_20250101"]), ["基本情報_20241003"])
        self.assertEqual(find_backup_sheets([
            "変更履歴", "介護サービス状況調査表", "20170113要介護度別状況調査表",
            "20161207要介護度別状況調査表", "20161108要介護度別状況調査表", "20161012要介護度別状況調査表",
        ]), ["20161207要介護度別状況調査表", "20161108要介護度別状況調査表", "20161012要介護度別状況調査表"])
        # 不同写法的日期也按日期先后比较
        self.assertEqual(find_backup_sheets(["一覧(2018-02-26)", "一覧_20190301"]), ["一覧(2018-02-26)"])
        # 名称不同的不互相影响
        self.assertEqual(find_backup_sheets(["A_20240101", "B_20230101"]), [])

    def test_separated_and_bracketed_dates_match_original(self):
        names = [
            "概要", "概要(2018-02-26)", "概要_2018.02.27", "概要（2018-02-28）", "概要 2018/03/01",
            "一覧", "一覧_2018-0226", "一覧_2018-02-30",
        ]
        # 分隔符前后不一致、无效日期都不算
        self.assertEqual(find_backup_sheets(names), names[1:5])

    def test_backup_with_original_is_found(self):
        # SSKJG2100_利用者負担額調整.xlsx
        self.assertEqual(find_backup_sheets([
            "変更履歴", "利用者負担額調整", "Sheet1", "利用者負担額調整_bak",
        ]), ["利用者負担額調整_bak"])

    def test_original_matches_ignoring_spaces_versions_and_width(self):
        # SSOPG1000_計画 実績設定.xlsx：原表名带空格，备份带版本号
        self.assertEqual(find_backup_sheets([
            "計画 実績設定", "20181213計画実績設定_v2.4", "20180529計画実績設定",
        ]), ["20181213計画実績設定_v2.4", "20180529計画実績設定"])
        self.assertEqual(find_backup_sheets(["ＡＢＣ一覧", "abc一覧_BAK"]), ["abc一覧_BAK"])

    def test_every_marker_requires_its_original(self):
        names = [
            "機能設計", "機能設計（バックアップ）", "旧機能設計", "Copy of 機能設計", "機能設計_old",
            "一覧", "一覧 (2)", "Sheet1のコピー", "20250101",
        ]
        self.assertEqual(find_backup_sheets(names), [
            "機能設計（バックアップ）", "旧機能設計", "Copy of 機能設計", "機能設計_old", "一覧 (2)",
        ])

    def test_copy_suffix_variants_require_original(self):
        names = [
            "変更履歴", "個別支援計画書_様式１（基本情報）",
            "個別支援計画書_様式１（基本情報）(2)_20221005", "個別支援計画書_様式１（基本情報）(2)",
            "個別支援計画書_様式１（基本情報）（３）", "一覧(2)",
        ]
        # 「一覧(2)」没有原表「一覧」，保留
        self.assertEqual(find_backup_sheets(names), names[2:5])

    def test_mmdd_suffix_requires_valid_date_and_original(self):
        names = [
            "変更履歴",
            "レイアウト(介護予防・日常生活支援総合事業費請求書1)", "レイアウト(介護予防・日常生活支援総合事業費請求書1)0513",
            "レイアウト(介護予防・日常生活支援総合事業費請求書2)", "レイアウト(介護予防・日常生活支援総合事業費請求書2)0513",
            "一覧", "一覧_0229", "一覧1340",
            "Sheet1", "Sheet10513",
            "帳票0513",
        ]
        # 1340 不是有效月日；Sheet10513 前面紧挨数字，分不清是不是月日；帳票0513 没有原表
        self.assertEqual(find_backup_sheets(names), [names[2], names[4], names[6]])

    def test_mmdd_does_not_join_dated_version_comparison(self):
        # 没有年份无法比较新旧，没有原表时都保留
        self.assertEqual(find_backup_sheets(["一覧0513", "一覧0601"]), [])


class CleanTargetTests(unittest.TestCase):
    def test_auto_generated_sheets_are_cleaned_without_original(self):
        names = [
            "変更履歴", "概要", "概要20220816", "自動生成_定数", "自動生成_OUTPUT_V_LIST_HEADER",
            "❓自動生成_JSON", "自动生成_引数", "帳票_自動生成ルール",
        ]
        self.assertEqual(find_clean_targets(names), [
            "概要20220816", "自動生成_定数", "自動生成_OUTPUT_V_LIST_HEADER", "❓自動生成_JSON", "自动生成_引数",
        ])
        # 只影响精简 Sheet，不算备份表（导出 PDF 时照常导出）
        self.assertEqual(find_backup_sheets(names), ["概要20220816"])


class ContainedSheetTests(unittest.TestCase):
    def test_name_containing_another_defaults_to_uncertain(self):
        names = [
            "一覧", "一覧詳細", "請求書", "請求書明細", "請求書_20240101",
            "Sheet1", "Sheet12", "表", "表紙", "❓概要", "概要メモ", "自動生成_一覧",
        ]
        # 已带 ❓ 的表也可作为被包含的名称（概要メモ）；已默认删除（請求書_20240101）、已带 ❓、自動生成、数字被切开（Sheet12）、单字被包含（表紙）都不算
        self.assertEqual(find_contained_sheets(names), ["一覧詳細", "請求書明細", "概要メモ"])

    def test_backup_targets_are_not_used_as_reference(self):
        names = ["概要", "概要_bak", "概要_bak詳細"]
        self.assertEqual(find_contained_sheets(names), ["概要_bak詳細"])
        self.assertEqual(find_contained_sheets(["一覧1340", "一覧"]), ["一覧1340"])


if __name__ == "__main__":
    unittest.main()
