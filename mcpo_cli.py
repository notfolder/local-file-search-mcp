import os
import sys
import json
import mcpo

def resource_path(relative_path):
    """実行ファイルまたはスクリプトからの相対パスを解決する"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstallerでビルドされた実行ファイル内の場合
        base_path = sys._MEIPASS
    else:
        # 通常のPythonスクリプト実行時
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def main():
    # コマンドライン引数をそのままmcpo.mainに渡す
    argv = sys.argv[1:]
    # --port 8000 --host 127.0.0.1 --config ./config.json --api-key "top-secret"
    default_args = {
        "--port": "8000",
        "--host": "127.0.0.1",
        "--config": "./config.json",
        "--api-key": 'top-secret',
    }
    for arg, value in default_args.items():
        if arg not in argv:
            argv.append(arg)
            argv.append(value)
    args_dict = {k.lstrip('-').replace('-', '_'): v if not v.startswith('-') else True
                for k, v in zip(argv, argv[1:] + ['--'])
                if k.startswith('-')}
    
    # configキーをconfig_pathに変換
    if 'config' in args_dict:
        config_path = resource_path(args_dict.pop('config'))
        args_dict['config_path'] = config_path
        
        # configファイルを読み込む
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        # mcpServers内のsearch_local_filesのargs配列を変換
        if 'mcpServers' in config:
            for server in config['mcpServers']:
                if 'search_local_files' in server:
                    if 'args' in server['search_local_files']:
                        server['search_local_files']['args'] = [
                            resource_path(arg) if arg == 'main.py' else arg
                            for arg in server['search_local_files']['args']
                        ]
        
        # 変更した設定を保存
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2)

    # portについては、int型に変換
    if 'port' in args_dict:
        args_dict['port'] = int(args_dict['port'])
    
    # mcpo パッケージのエントリポイントを呼び出し
    mcpo.main(**args_dict)

if __name__ == "__main__":
    main()
