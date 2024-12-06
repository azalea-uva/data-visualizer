import os
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.cm import jet
import matplotlib as mpl
# グラフの全体的なフォントサイズを設定
mpl.rcParams['font.size'] = 12

# フォントをArialに設定
plt.rcParams['font.family'] = 'Arial'

cmap = jet

# グラフのサイズを黄金比に設定
golden_ratio = (1 + 5 ** 0.5) / 2
fig_width = 8  # 幅を8インチに設定
fig_height = fig_width / golden_ratio

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
        plt.figure(figsize=(fig_width, fig_height))
        for n in range(1, 7):
            x_column = df.columns[4 * n - 3]
            y_column = df.columns[4 * n - 4]

            if polarity == 'p':
                # 極性がpの場合、y軸の値に-1を掛ける
                df[y_column] = -df[y_column]
                # 凡例の名前を設定（n=1から6の範囲で-20(n-1)）
                if n == 1:
                    legend_name = f'$V_G$ = {-20 * (n - 1)} V'
                else:
                    legend_name = f'{-20 * (n - 1)} V'
            else:
                if n == 1:
                    legend_name = f'$V_G$ = {20 * (n - 1)} V'
                else:
                    legend_name = f'{20 * (n - 1)} V'
                
            color = cmap(n / len(y_column)) 
            plt.plot(df[x_column], df[y_column], color=color, label=legend_name)
        
        plt.xlabel(r"V$_{DS}$ (V)", fontsize="x-large")
        plt.ylabel(r"|I$_{DS}$| (A)", fontsize="x-large")
        plt.legend()
        plt.xlim(-100, 5)
        plt.ylim(df[y_column].min()/50, )
        # plt.title("output")
        # plt.figure(figsize=(fig_width, fig_height))
        
        # グラフを300dpiのJPEG形式で保存（カレントディレクトリ内に保存）
        output_filename = os.path.join(directory_path, "output_graph.jpg")
        plt.savefig(output_filename, dpi=300, format='jpeg')

        plt.show()
    else:
        print("無効なファイル番号が選択されました。")
