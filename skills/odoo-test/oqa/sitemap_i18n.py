"""English/Japanese text used by the sitemap CLI and MCP responses.

Odoo labels come from Odoo itself; these templates only translate our explanations.
Technical names, selectors and expressions are always passed through unchanged.
"""

JA = {
    "Sitemap languages must be en_US / ja_JP (or en / ja).": "サイトマップの言語には en_US / ja_JP（または en / ja）を指定してください。",
    "Activate these languages in Odoo before building the sitemap: {languages}": "サイトマップを作成する前に、Odoo で次の言語を有効にしてください: {languages}",
    "Activate English (en_US) or Japanese (ja_JP) in Odoo before building the sitemap.": "サイトマップを作成する前に、Odoo で英語（en_US）または日本語（ja_JP）を有効にしてください。",
    "not readable with this profile, so menus and actions have no external id and `%(module.name)d` buttons cannot be followed: {error}": "このプロファイルでは外部 ID を読み取れないため、メニューとアクションの外部 ID を取得できず、`%(module.name)d` ボタンの遷移先を追跡できません: {error}",
    "the views of {model} could not be read: {error}": "{model} のビューを読み取れませんでした: {error}",
    "the view definition is not readable XML": "ビュー定義を XML として読み取れません",
    "button '{label}' points at action '{name}', which is not a known action": "ボタン「{label}」が参照するアクション '{name}' は不明です",
    "button '{label}' points at action number {number}, which could not be read": "ボタン「{label}」が参照するアクション番号 {number} を読み取れませんでした",
    "leads to {target}, which could not be read": "遷移先 {target} を読み取れませんでした",
    "No sitemap for this workspace yet: run `oqa sitemap build` (it reads the menus and views, it clicks nothing).": "このワークスペースにはサイトマップがありません。`oqa sitemap build` を実行してください（メニューとビューの読み取りのみで、クリックは行いません）。",
    "This sitemap was written by a newer odoo-test: update the skill.": "このサイトマップは新しいバージョンの odoo-test で作成されています。スキルを更新してください。",
    "sitemap of {profile} on Odoo {version}, built {date}: {menus} menus, {actions} actions, {views} views": "Odoo {version} のサイトマップ（プロファイル: {profile}、作成日: {date}）：メニュー {menus} 件、アクション {actions} 件、ビュー {views} 件",
    ". It is {days} days old: if a module was installed since, run `oqa sitemap build` again.": "。作成から {days} 日経過しています。モジュールを追加した場合は `oqa sitemap build` を再実行してください。",
    ". Not complete: the crawl stopped at its size limit.": "。サイズ上限に達したため、サイトマップは未完成です。",
    ". Label languages: {languages}": "。ラベルの言語: {languages}",
    ". Languages not active in Odoo: {languages}": "。Odoo で未有効の言語: {languages}",
    "menu {label}": "メニュー {label}",
    '{type} action "{label}"': '{type} アクション「{label}」',
    " on {model}": "（モデル: {model}）",
    " (a dialog)": "（ダイアログ）",
    "{type} view of {model}": "{model} の {type} ビュー",
    "button '{label}'": "ボタン「{label}」",
    "field '{label}' [name={name}]": "フィールド「{label}」[name={name}]",
    "    ... {count} more not shown: narrow with {hint}": "    … 他 {count} 件は省略しています。{hint} で絞り込んでください。",
    "  NOT possible here: {disabled}": "  このビューでは使用不可: {disabled}",
    "nothing matches '{text}'{scope}. Try fewer or other words (a menu name, a model, a button label).": "「{text}」に一致する項目がありません{scope}。語を減らすか、別のメニュー名・モデル名・ボタン名で検索してください。",
    " among {kind}s": "（対象: {kind}）",
    "nothing in the sitemap matches '{text}'": "サイトマップに「{text}」と一致する項目がありません",
    "-> unknown until it is run (it calls a method)": "-> 実行するまで遷移先は不明（メソッド呼び出し）",
    "-> not resolved": "-> 遷移先を解決できません",
    '    {stat}button "{label}" [name={name}] {selector}': '    {stat}ボタン「{label}」[name={name}] {selector}',
    "stat ": "統計 ",
    " shown only if: not ({condition})": " 表示条件: not ({condition})",
    "    clicking a row -> {target}": "    行をクリック -> {target}",
    "required": "必須",
    "readonly": "読み取り専用",
    "{flag} if: {condition}": "{flag} の条件: {condition}",
    '    {name} "{label}" {type}': '    {name}「{label}」{type}',
    "  options: {options}": "  選択肢: {options}",
    "no way found to {target} from {source}: it may be reached only by a button that calls a method.": "{source} から {target} への経路が見つかりません。メソッドを呼び出すボタンからのみ到達できる可能性があります。",
    "any menu": "いずれかのメニュー",
    "1. start at {target}": "1. 開始: {target}",
    "open the submenu": "サブメニューを開く",
    "it opens": "開く",
    "it shows": "表示する",
    "click a row": "行をクリック",
    'click the button "{label}" {selector}': 'ボタン「{label}」をクリック {selector}',
    "Sitemap written to {file}": "サイトマップの保存先: {file}",
    "{buttons} buttons call a method (where they lead is only known by running them); {unresolved} things could not be read.": "メソッドを呼び出すボタン: {buttons} 件（遷移先は実行時にのみ判明）。読み取りできなかった項目: {unresolved} 件。",
}


def text(language, template, **values):
    return (JA.get(template, template) if language == "ja" else template).format(**values)
