# LinkSphere パラメーターシート ver0.7 変更点

- 元ファイル: `LinkSphere_パラメーターシート_ver0.6_スタック実機反映.xlsx`
- 新ファイル: `LinkSphere_パラメーターシート_ver0.7_Huawei準拠.xlsx`
- 根拠: Huawei CloudEngine S5735-L-V2 の V600R025C00 公式マニュアル（英語版。V600 の日本語版はありません）

## 作り方

元の機種シート 6 枚は、行を足したときにセル結合がずれて崩れていました。手で直すと同じ崩れが残るため、6 枚とも同じ枠組みで作り直しました。「改訂履歴」シートは元のまま引き継ぎ、ver0.7 の行を足しました。

生成スクリプトは `output_system/param_sheet/` にあります。値を直すときは `models.py` を書き換えてから、次のコマンドで作り直せます。

```bash
python3 output_system/param_sheet/build_param_sheet.py <元のxlsx> <出力xlsx>
```

## 直した箇所

### ファイルの崩れ・読めない箇所

| 指摘 | 直し方 |
|---|---|
| C48T4XE-L のセル結合が 1 行ずれ、見出しや値が隠れていた | 作り直し。ポート構成図も追加 |
| 行 22～24 の値が切れて読めなかった（「内蔵AC電源」→「内」） | 値の欄を広く取り直した |
| 10GE のポート図が 4 シートで「1/1/x」になっていた | 全シートを「1/0/x」に統一 |
| C24T4XE-L の製品名の欄に「LinkSphere C24-L」があった | 見出しのない欄ごと削除 |
| C24P4XE-L に「Cf」が残っていた | 削除 |
| 印刷範囲がなく、130 ページに分かれていた | A3 縦・横 1 ページに合わせた。全体で 27 ページ |
| 欄外のメモが印刷で切れていた | 各ブロックの「備考」欄へ移した |
| C8P4X-Q-L・C24T4XE-L だけ見出しが違っていた | 全シートで同じ見出しにそろえた |

### Huawei の書き方に合わせた箇所

| 元の記載（Cisco） | ver0.7（Huawei） | 根拠 |
|---|---|---|
| ポート図の「Gi」 | GE（10GE ポートは 10GE） | [Port Numbering Conventions](https://support.huawei.com/enterprise/en/doc/EDOC1100277349/884016b/port-numbering-conventions)「the first GE port is numbered GE1/0/1」 |
| VTP 設定の節 | 削除 | VTP は Cisco 独自のプロトコル |
| Mode「Switch Port」 | link-type（既定 negotiation-auto） | [Default Settings for VLANs](https://support.huawei.com/enterprise/en/doc/EDOC1100515500/22c6384b/default-settings-for-vlans)「S5735-L-V2 … negotiation-auto」 |
| PortFast | エッジポート（stp edged-port enable） | [Configuring Edge Ports](https://support.huawei.com/enterprise/en/doc/EDOC1100515500/f8fc09d6/configuring-edge-ports-and-bpdu-filtering-ports-for-rstp-mstp) |
| ポートごとの BPDU Guard | 削除。BPDU 保護はシステム全体の設定（既定は無効） | [STP Commands](https://support.huawei.com/enterprise/en/doc/EDOC1100515397/81575409/stp-rstp-mstp-configuration-commands)「By default, BPDU protection is disabled」 |
| Loop Guard | ループ保護（stp loop-protection） | [Configuring Loop Protection](https://support.huawei.com/enterprise/en/doc/EDOC1100515500/dfe988c5/configuring-loop-protection) |
| ポートチャネル設定 | Eth-Trunk 設定（既定 manual、LACP は lacp-static） | [Creating an Eth-Trunk](https://support.huawei.com/enterprise/en/doc/EDOC1100515500/d542520f/creating-an-eth-trunk-interface-and-configure-a-link-aggregation-mode-for-it) |
| Eth-Trunk ごとの load-balance | 全体設定の load-balance profile（Eth-Trunk ごとの設定は S5735-L-V2 非対応） | [Static Load Balancing](https://support.huawei.com/enterprise/en/doc/EDOC1100515500/480e946/configuring-a-static-load-balancing-mode) |
| 静的ルートの Distance | Preference（既定 60） | [Static Route Preference](https://support.huawei.com/enterprise/en/doc/EDOC1100515486/b22aff1d/optional-setting-a-default-preference-value-for-ipv4-static-routes) |
| VRRP の Hello/Hold | 通知間隔（既定 1 秒）・プリエンプト遅延 | [VRRP Advertisement Interval](https://support.huawei.com/enterprise/en/doc/EDOC1100515472/130bb8b8/configuring-an-interval-for-sending-vrrp-advertisement-packets) |
| Track Interface/Priority Decre | 監視IF／優先度減算（既定 10） | [VRRP Interface Tracking](https://support.huawei.com/enterprise/en/doc/EDOC1100515472/563aeaac/associating-vrrp-with-interface-status-to-trigger-master-backup-vrrp-switchovers) |
| Enable Password、SSH のドメイン名の備考 | 削除。SSH の受付インターフェースと STelnet の欄を追加 | [SSH Commands](https://support.huawei.com/enterprise/en/doc/EDOC1100515397/7fa11d5c/ssh-configuration-commands)「an SSH server does not accept login requests from any interface by default」 |
| Passive Interface／Redistribute | silent-interface／import-route、network はワイルドカードマスク | [OSPF Commands](https://support.huawei.com/enterprise/en/doc/EDOC1100515397/758505a5/ospf-configuration-commands) |
| Storm Control | storm control（action は error-down／block／suppress） | [Configuring Storm Control](https://support.huawei.com/enterprise/en/doc/EDOC1100515473/6d65ae1f/configuring-storm-control) |
| キュースケジューリング | PQ／DRR（qos schedule-profile で設定） | [Congestion Management](https://support.huawei.com/enterprise/en/doc/EDOC1100515480/c91a7116/configuring-congestion-management) |

### 足した欄

| 欄 | 中身 | 根拠 |
|---|---|---|
| スタック構成情報 | メンバーID・優先度・スタックポート・スタックドメイン・DAD | [Stack Commands](https://support.huawei.com/enterprise/en/doc/EDOC1100515397/ec3bf9b2/stack-configuration-commands)、[Configuring DAD](https://support.huawei.com/enterprise/en/doc/EDOC1100512818/170e6a23/configuring-dad) |
| メンバー 2～4 のポート設定 | GE2/0/x～GE4/0/x（先頭の数字＝メンバーID） | [Stack Login and Management](https://support.huawei.com/enterprise/en/doc/EDOC1100512818/abe1783/stack-login-and-management)「if the stack ID is 2, the interface number changes to 2/0/1」 |
| SNMPv3 ユーザー | 認証方式・暗号方式・セキュリティレベル | [Basic SNMPv3](https://support.huawei.com/enterprise/en/doc/EDOC1100515474/d29954d9/configuring-basic-snmpv3-functions) |
| PoE 設定（PoE 機種だけ） | 給電（既定 有効）・優先度（既定 low）・上限電力（既定 30000 mW） | [PoE Commands](https://support.huawei.com/enterprise/en/doc/EDOC1100515397/8ba752ac/poe-configuration-commands) |
| MST リージョン名 | 既定は管理IFの MAC アドレス | [STP Commands](https://support.huawei.com/enterprise/en/doc/EDOC1100515397/81575409/stp-rstp-mstp-configuration-commands) |
| 管理用 Vlanif | VLAN ID・IP・マスク・デフォルトゲートウェイ | C48T4XE-L にあった方針（物理管理ポートではない）を全シートへ展開 |

空欄だった既定値も埋めました: LLDP は有効（[Default Settings for LLDP](https://support.huawei.com/enterprise/en/doc/EDOC1100515474/e4aee5ae/default-settings-for-lldp)）、ポートの Speed／Duplex は GE が Auto・10GE が 10G／Full（[Auto-Negotiation](https://support.huawei.com/enterprise/en/doc/EDOC1100515479/6b31738a/configuring-an-interface-to-work-in-auto-negotiation-mode)）。

## 記入・確認が要る欄（黄色のセル）

| シート | 欄 | 理由 |
|---|---|---|
| C8T4X-Q-L・C8P4X-Q-L | スタックポート 1・2（メンバーごと） | 専用スタックポートがなく、どの 10GE ポートを使うかは設計で決める |
| C8P4X-Q-L | 電源構成・PoE 対応規格・PoE 給電能力・スタックポートの有無 | 元のシートでも空欄。データシートの値が要る |
| C24T4XE-L | 電源構成・冷却方式 | 元のシートは稼働状態（PWR1／FAN1 正常）だけで、仕様が書かれていない |
| 全シート | DAD（direct／relay） | 構成で決める |
| 全シート | SSH 受付インターフェース | 管理用 Vlanif を決めてから記入する |

## 判断を残した点

- **改訂履歴**: 既存の「第 2 版／YYYY/MM/DD／新規」の行は、中身が分からないため書き換えていません。ver0.1～0.6 の履歴、改訂者、承認者の記入をお願いします。
- **OEM 元の型番・バージョン**（V600R025C00SPC500、S5735-V2_…cc）は、「Huawei と同じでよい」との指示どおり残しました。顧客へ出す前に扱いを決めてください。
- **シートの中の値は既定値**: 各シートの値は機器の初期値です。案件ごとの設計値は別途記入が要ります。
- **マニュアルで確定できなかった点**:
  - SSH の鍵生成にドメイン名が要るか（マニュアルの前提条件に記載なし）
  - コンソール認証が初回ログイン後に password と aaa のどちらになるか（シートは実機確認値の password のまま）
  - MEth のない機種での MST リージョン名の既定値
