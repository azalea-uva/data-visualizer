import os
import pandas as pd
import matplotlib.pyplot as plt

# Pythonファイルが現在実行されているディレクトリを取得
directory_path = os.path.dirname(os.path.abspath(__file__))

# 指定したディレクトリ内のxlsファイルのリストを取得
xls_files = [f for f in os.listdir(directory_path) if f.endswith('.xls')]

if not xls_files:
    print("ディレクトリ内にxlsファイルが見つかりませんでした。")
else:
    print("利用可能なxlsファイル:")
    for i, xls_file in enumerate(xls_files):
        print(f"{i + 1}: {xls_file}")

    # ユーザーが選択したファイルのインデックスを入力
    selected_index = int(input("プロットするファイルの番号を選択してください: ")) - 1

    if 0 <= selected_index < len(xls_files):
        selected_file = xls_files[selected_index]

        # pかnかを選択
        polarity = input("プロットするデータの極性を選択してください（pまたはn）: ")

        # xlsファイルを読み込む
        xls_path = os.path.join(directory_path, selected_file)
        df = pd.read_excel(xls_path)

        # プロットを設定
        plt.figure()
        for n in range(1, 7):
            x_column = df.columns[4 * n - 3]
            y_column = df.columns[4 * n - 4]

            if polarity == 'p':
                # 極性がpの場合、y軸の値に-1を掛ける
                df[y_column] = -df[y_column]
                # 凡例の名前を設定（n=1から6の範囲で-20(n-1)）
                legend_name = f'n={20 * (n - 1)}'
            else:
                # 凡例の名前を設定（n=1から6の範囲で20(n-1)）
                legend_name = f'n={-20 * (n - 1)}'

            plt.plot(df[x_column], df[y_column], label=legend_name)

        plt.xlabel("V_DS (V)")
        plt.ylabel("|I_DS| (A)")
        plt.legend()
        plt.title("output")

        # グラフを300dpiのJPEG形式で保存（カレントディレクトリ内に保存）
        output_filename = os.path.join(directory_path, "output_graph.jpg")
        plt.savefig(output_filename, dpi=300, format='jpeg')

        plt.show()
    else:
        print("無効なファイル番号が選択されました。")
