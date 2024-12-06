import os
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from scipy.stats import linregress

# スクリプトが保存されているディレクトリのパスを取得
script_directory = os.path.dirname(os.path.abspath(__file__))

# スクリプトが保存されているディレクトリ内のxlsファイルをリストアップ
xls_files = [file for file in os.listdir(script_directory) if file.endswith('.xls')]

# エクセルファイルが見つからない場合
if not xls_files:
    print("ディレクトリ内にxlsファイルが見つかりませんでした。")
    exit()

# CSVファイルに保存するためのデータを格納するリスト
csv_data = []

# 最初のxlsファイルのx軸の範囲を保持する変数
x_min_global = None
x_max_global = None

# すべてのxlsファイルに対して処理を繰り返す
for selected_excel_file in xls_files:
    # 選択されたエクセルファイルのパスを作成
    selected_excel_path = os.path.join(script_directory, selected_excel_file)

    # 選択されたエクセルファイルを読み込む
    try:
        df = pd.read_excel(selected_excel_path, engine='xlrd')

        # 4列目をx，1列目をyとしてグラフにプロット
        x = df.iloc[:, 3]  # 4列目
        y = df.iloc[:, 0]  # 1列目

        # yの値が負である場合、絶対値を取得
        y = y.abs()

        # yの平方根を計算
        y_sqrt = np.sqrt(y)

        # グラフの形状を3:2に設定
        fig_width = 6.0
        fig_height = 4.0
        fig, ax1 = plt.subplots(figsize=(fig_width, fig_height))
        
        ax1.plot(x, y, label='|I_DS| (A)', color='tab:blue')
        ax1.set_xlabel('V_G (V)')
        ax1.set_ylabel('|I_DS| (A)', color='tab:blue')
        ax1.set_yscale('log')  # 左軸を対数スケールに設定
        
        # エクセルファイル名の最初の2文字をグラフタイトルに設定
        graph_title = selected_excel_file[:2]
        ax1.set_title(f'{graph_title}')
        
        ax2 = ax1.twinx()
        ax2.plot(x, y_sqrt, label='sqrt(|I_DS|) (A^1/2)', color='tab:orange', linestyle='--')  # 破線に設定
        ax2.set_ylabel('sqrt(|I_DS|) (A^1/2)', color='tab:orange')
        ax2.set_yscale('linear')  # 右軸を線形スケールに設定

        # 凡例をグラフの左外側に表示
        lines1, labels1 = ax1.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        lines = lines1 + lines2
        labels = labels1 + labels2
        ax2.legend(lines, labels, loc='upper right', bbox_to_anchor=(-0.15, 0.9))  # 凡例を左側に表示
            
        # グラフの余白を調整して軸ラベルが見切れないようにする
        fig.tight_layout()
        
        # エクセルファイル名を取得し、同じ名前でグラフを保存
        excel_file_name = os.path.splitext(selected_excel_file)[0]  # 拡張子を除いたエクセルファイル名
        graph_file_name = f"{excel_file_name}_graph.jpg"
        graph_file_path = os.path.join(script_directory, graph_file_name)
        plt.savefig(graph_file_path, dpi=300, format='jpeg')

        #plt.show()

        # 最初のxlsファイルの場合、x軸の範囲をユーザーに尋ねて計算
        if x_min_global is None:
            plt.show()
            x_range = input("x軸の範囲を指定してください (例: 0.5 1.5): ").split()
            if len(x_range) == 2:
                x_min_global, x_max_global = map(float, x_range)

        # すでにx軸の範囲が計算されている場合、それを使用
        if x_min_global is not None and x_max_global is not None:
            mask = (x >= x_min_global) & (x <= x_max_global)
            x_subset = x[mask]
            y_sqrt_subset = y_sqrt[mask]
            slope, intercept, _, _, _ = linregress(x_subset, y_sqrt_subset)
            approx_line = slope * x + intercept
            V_th = -intercept / slope
            L = 50 #um チャネル長
            W = 1000 #um 電極幅
            Ci = 1e-4 #F/m^2 単位面積当たりの誘電率
            mu = (2 * L * slope ** 2 * 10000) / (Ci * W) 
            
            # 二度目のグラフを作成し、近似直線を追加して表示
            fig2, ax1 = plt.subplots(figsize=(fig_width, fig_height))
            ax1.plot(x, y, label='|I_DS| (A)', color='tab:blue')
            ax1.set_xlabel('V_G (V)')
            ax1.set_ylabel('|I_DS| (A)', color='tab:blue')
            ax1.set_yscale('log')  # 左軸を対数スケールに設定
            ax1.set_title(f'{graph_title}')
            
            ax2 = ax1.twinx()
            ax2.plot(x, y_sqrt, label='sqrt(|I_DS|) (A^1/2)', color='tab:orange', linestyle='--')  # 破線に設定
            ax2.set_ylabel('sqrt(|I_DS|) (A^1/2)', color='tab:orange')
            ax2.set_yscale('linear')  # 右軸を線形スケールに設定
            ax2.plot(x, approx_line, label='Approx Line', color='tab:red', linestyle=':')  # 近似直線を追加
            
            # 凡例をグラフの左外側に表示
            lines1, labels1 = ax1.get_legend_handles_labels()
            lines2, labels2 = ax2.get_legend_handles_labels()
            lines = lines1 + lines2
            labels = labels1 + labels2
            ax2.legend(lines, labels, loc='upper right', bbox_to_anchor=(-0.15, 0.9))  # 凡例を左側に表示

            # グラフの余白を調整して軸ラベルが見切れないようにする
            fig2.tight_layout()
            
            # 追加のグラフを保存
            approx_graph_file_name = f"{excel_file_name}_approx_graph.jpg"
            approx_graph_file_path = os.path.join(script_directory, approx_graph_file_name)
            plt.savefig(approx_graph_file_path, dpi=300, format='jpeg', bbox_inches='tight')  # bbox_inchesを設定して余白を削除

            #plt.show()

            # yのデータの最小値と最大値を出力
            y_min = y.min()
            y_max = y.max()
            y_ratio = y_max / y_min

            #print(f"yのデータの最小値: {y_min}")
            #print(f"yのデータの最大値: {y_max}")
            #print(f"最大値と最小値の比: {y_ratio}")

            #print(f"グラフを {approx_graph_file_name} として保存しました。")
            
            # データをCSVに追加
            csv_data.append([excel_file_name, V_th, mu, y_min, y_max, y_ratio])

        else:
            print("無効な入力です。x軸の範囲を正しく指定してください。")
        
    except Exception as e:
        print(f"エクセルファイルの読み込み中にエラーが発生しました: {str(e)}")

# CSVファイルにデータを保存
csv_file_path = os.path.join(script_directory, "data_summary.csv")
csv_columns = ["ファイル名", "閾値電圧 (V)", "移動度 (cm^2/V s)", "I_off", "I_on", "I_on/I_off"]
df_summary = pd.DataFrame(csv_data, columns=csv_columns)

# 平均値と標準偏差を計算して追加
average_values = df_summary.mean()
std_deviation = df_summary.std()
df_summary = df_summary.append(average_values, ignore_index=True)
df_summary = df_summary.append(std_deviation, ignore_index=True)
df_summary.iloc[-2, 0] = '平均値'
df_summary.iloc[-1, 0] = '標準偏差'

df_summary.to_csv(csv_file_path, index=False, encoding='utf-8-sig')  # UTF-8エンコーディングで保存
print(f"データのサマリーを {csv_file_path} に保存しました。")
