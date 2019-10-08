import os
import glob
import time

def execute_logic(file_path, progress_bar, after_complete_process):
    """
    GUIから実行されるアプリケーションのロジック本体。
    インプットファイルに対しての処理を記述する。
    """
    ### アプリケーションの実装 ###
    for i in range(10):
        time.sleep(1)
        progress_bar.configure(val=(i+1)*10)

    # 実行完了時の処理
    after_complete_process()

def get_file_path_in_folder(folder_path, search_extension):
    """
    フォルダ内を再帰的に調べて、指定した拡張子を持つファイルの
    パスをすべて取得する。

    Parameters
    ----------
    folder_path : str
        検索を行うフォルダのパス
    search_extension : str
        パスを取得するファイルの拡張子
    """

    search_extension = '*.' + search_extension
    return glob.glob(os.path.join(folder_path, '**', search_extension),
                     recursive=True)