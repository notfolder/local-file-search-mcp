import sys
import mcpo
def main():
    # コマンドライン引数をそのままmcpo.mainに渡す
    argv = sys.argv[1:]
    # --port 8000 --host 127.0.0.1 --config ./config.json --api-key "top-secret"
    default_args = {
        "--port": "8000",
        "--host": "127.0.0.1",
        "--config": "./config.json",
        "--api-key": '"top-secret"',
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
        args_dict['config_path'] = args_dict.pop('config')
    # portについては、int型に変換
    if 'port' in args_dict:
        args_dict['port'] = int(args_dict['port'])
    
    # mcpo パッケージのエントリポイントを呼び出し
    mcpo.main(**args_dict)

if __name__ == "__main__":
    main()
