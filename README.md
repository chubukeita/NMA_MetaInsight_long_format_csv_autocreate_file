# NMA_MetaInsight_long_format_csv_autocreate_file
Network Meta-analysisを実施するためのソフトウェアにMetaInsightがある。これを利用するためには、Long format形式か、Wide format形式のcsvファイルをアップロードする必要がある。ただし、データを手入力してこれらの形式のcsvを、対象のアウトカムの数だけ作成することは非常に時間がかかり、ヒューマンエラーを誘発する原因となりえる。よって、より直感的でデータの入力がしやすいフォーマットを作成し、入力データに基づいてLong format形式のcsvファイルに変換するマクロ有効ブックを、ChatGPT 4.0によるリファクタリングを経て実装した（以下，入力ツール）。入力ツールによりLong format形式のcsvをボタン操作で作成できるようにした。
