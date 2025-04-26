openwebuiと組み合わせて、**ローカルファイルをアップロードせずにLLMに読んでもらえる**仕組みができました！

⚠️ **注意**：プロンプトにローカルファイル情報が出るため、**セキュリティーには十分注意**してください。
ローカルや社内AIを使わない場合、ローカルファイル情報がダダ漏れになるリスクがあります。この記事で紹介するコードを使用したことによる損害について、筆者は責任を負いかねますのでご了承ください。

---
## 何ができるか
「ローカルファイルからXXXDBの仕様書を探して読んで、XXXテーブルからXXX条件でデータを抽出するsql文を考えて」ができる様になります。


## 概要
- **MCPサーバー**を作成（GitHub公開）
- **openwebui**と連携して、ファイル検索をLLMに頼める
- ファイルアップロード不要。ローカルインデックス活用


## 手順

### 1. openwebuiをDockerで起動
まずはopenwebuiをDockerで立ち上げます。

```bash
docker run -d -p 8080:8080 --add-host=host.docker.internal:host-gateway -v open-webui:/app/backend/data --name open-webui ghcr.io/open-webui/open-webui:main
```

初期設定（APIキー設定など）は各自で行ってください。ChatGPT APIやローカルLLMと接続できます。


### 2. 拙作MCPサーバーをmcpo経由で起動
次にMCPサーバーを立ち上げます。

```bash
git clone https://github.com/notfolder/local-file-search-mcp.git
cd local-file-search-mcp
conda create -f env.yml
conda activate mcp-file-search-env
mcpo --port 8000 --host 127.0.0.1 --config ./config.json --api-key "top-secret"
```

下記コマンドでアプリ化中。まだうごきません。。。
```bash
$ pyinstaller mcpo_cli.py main.py --add-data config.json:. --add-data main.py:. --onefile --console -n local-file-search-mcp
```
⚠️プロセスがkillできなくなるので注意！
※listenするポートやbindするホスト名はmcpo_cli.pyの内容を変更して下さい。


### 3. openwebuiにMCPサーバーを設定
- 「管理者設定」→「ツール」→「＋」をクリック
- 以下を設定します：
  - URL: `http://host.docker.internal:8000/search_local_files`
  - Auth: Bearerとして`top-secret`を指定
  - Visibility: `Public`


### 4. システムプロンプトに指示を追加
openwebuiの**個人設定**→**一般**→**システムプロンプト**に以下を追記してください：

> 「ローカルファイルに関する検索や読み込みを依頼されたら、search_local_filesを利用するようにして下さい。」


### 5. 利用開始！
チャットを開き、「＋」ボタンで`search_local_files`ツールを有効にします。

例：
```text
ローカルファイルで内容に「FastMCP」が含まれるファイルを探して
```

ファイルが見つかったら、さらに：
```text
「検索結果のファイル」を読んで、概要を教えて
```

textファイルだけでなく、Word・Excel・PowerPointファイルにも対応しています。


---

## 番外編：開発の裏話
このツールの開発には**GitHub Copilot**をフル活用しました！

Windowsの`Search.CollatorDSO`や、Macの`NSMetadataQuery`なんて全然知らなかったので、頼りきりで作りました。
ちなみに、いまだに**WindowsのDocumentsフォルダ限定検索**の方法が見つかっていません……（誰か教えてください🙏）。

---

## リンク
- GitHubリポジトリ：[local-file-search-mcp](https://github.com/notfolder/local-file-search-mcp)


---

ここまで読んでいただき、ありがとうございました！
ローカルファイルをもっと安全に、便利に扱いたい方はぜひ試してみてください🚀

## この記事の作成過程
https://chatgpt.com/share/680cb0c0-dff4-8009-aca6-acc040c58dfd
