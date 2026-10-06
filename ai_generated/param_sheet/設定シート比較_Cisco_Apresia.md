# 設定シート比較（Cisco・Apresia と LinkSphere ver0.7）

LinkSphere パラメーターシート ver0.7 に、社内の Cisco と Apresia の設定シートにある大項目が欠けていないかを比べた結果です。ポート数など機種で当然変わる違いは対象外です。

## 比べた資料

| 略称 | 資料名 | 更新日 | 場所 |
|---|---|---|---|
| C9200L | 設定シート_Cisco Catalyst 9200L.xlsx | 2025-09-29 | SharePoint vendor_00053 / DocLib7 |
| C9300L | 設定シート_Cisco Catalyst 9300L.xlsx | 2025-09-26 | SharePoint vendor_00053 / DocLib7 |
| NP3000 | 設定シート_Apresia NP3000_Ver1.0.0.xlsx | 2026-09-30 | SharePoint topics_01343 / Library / Apresia |
| NP4000 | 設定シート_Apresia np4000／2000シリーズ.xlsx | 2023-10-20 | SharePoint topics_01343 / Library / Apresia |

いずれも案件用のコピーではなく、ひな形（テンプレート）です。LinkSphere ver0.6 は C9200L と同じ様式なので、C9200L から作られたと考えられます。

## 結論

- ver0.7 に無く、Cisco か Apresia の設定シートにある大項目は **13 個**です。
- 「表紙」と「システム設定（設計値）」は機能ではなく様式の問題です。どちらも足すことをお勧めします。
- 機能の 11 個のうち 9 個（DNS・MTU・RADIUS・DHCP・Loopback・Route-map・リング冗長・QinQ・Web認証）は、Huawei（S5735-L-V2、V600R025C00）で使えます。
- Netflow は Huawei 側が非対応です（代わりに sFlow が使えます）。ip device tracking は Huawei 側の相当機能を調べていません。
- 逆に ver0.7 にある「Routed Port設定」は、S5735-L-V2 では使えない見込みです。削除を検討してください。

記号: ○ = 大項目としてある、△ = 別の節の中に一部ある、× = 無い

## 1. 大項目の比較

### ver0.7 に無い大項目

| 大項目 | C9200L | C9300L | NP3000 | NP4000 | Huawei での対応 | お勧め |
|---|---|---|---|---|---|---|
| 表紙（文書番号・作成／審査／承認の欄） | ○ | ○ | ○ | ○ | － | **追加**。4 つとも持っている |
| DNS設定 | × | ○ | ○ | ○ | 対応（dns resolve、dns server） | **追加** |
| MTU設定 | ○ | ○ | ○ | ○ | 対応（jumboframe enable、mtu） | **追加**。4 つとも持っている |
| RADIUS設定（認証サーバー） | × | ○ | × | × | 対応（RADIUS／HWTACACS） | **追加** |
| DHCP設定（リレー） | × | × | ○ | △ | 対応（DHCP リレー・サーバー・スヌーピング） | **追加** |
| システム設定（設計値） | ○ | ○ | × | ○ | － | **追加**。ver0.7 は初期状態の一覧だけで、変える値を書く欄が無い |
| Loopback Interface設定 | × | ○ | × | × | 対応（interface loopback） | L3 で使うなら追加 |
| Route-map設定 | × | ○ | × | × | 対応（route-policy） | 動的ルーティングを使うなら追加 |
| リング冗長（Apresia の MMRP-Plus） | × | × | ○ | × | 対応（ERPS、SEP） | リング構成を扱うなら追加 |
| QinQ設定 | × | × | ○ | ○ | 対応（dot1q-tunnel、vlan-stacking） | 使う案件があれば追加 |
| Web認証 | × | × | △ | △ | 対応（802.1X・MAC 認証・Portal 認証） | 使う案件があれば追加 |
| Netflow設定 | × | ○ | × | × | NetStream は**非対応**。sFlow は対応 | 足すなら「sFlow設定」として追加 |
| ip device tracking設定 | × | ○ | × | × | 未調査 | 優先度は低い |

### 足さなくてよい大項目（Cisco・Apresia 固有）

| 大項目 | ある資料 | 理由 |
|---|---|---|
| VTP設定 | C9200L・C9300L・NP4000 | Cisco 独自のプロトコル。ver0.7 で削除済み |
| CDP設定 | C9200L・C9300L | Cisco 独自のプロトコル。ver0.7 の LLDP設定で代わりになる |
| Smart License設定 | C9200L・C9300L | Cisco のライセンス方式。Huawei は今回の機能すべてで追加ライセンスが不要 |
| Default-Gateway設定 | C9300L | ver0.7 は「管理用インターフェース」に Default Gateway の欄がある |

### 両方にある大項目（問題なし）

ハードウェア構成情報、スタック、管理アクセス設定、時刻設定、Syslog設定、SNMP設定、STP設定、LLDP設定、管理ポート、バナー設定、VLAN設定、SVI（ver0.7 は Vlanif設定）、ポート構成情報、ポート設定、ポートチャネル（ver0.7 は Eth-Trunk設定）、ルーティング設定、ACL設定、QoS設定、ストーム制御設定。

ver0.7 にだけある大項目は「PoE設定」です。

## 2. 既存の大項目の中で欠けている欄

| 大項目 | 欠けている欄 | ある資料 | Huawei での対応 |
|---|---|---|---|
| SNMP設定 | SNMPv2c のコミュニティ（Read Only／Read Write） | C9200L・C9300L・NP3000・NP4000 | 対応。ただし v1／v2c は WEAKEA（弱い暗号方式の追加パッケージ）が要る |
| SNMP設定 | 送る Trap の種類（Traps） | 4 つとも | 対応（snmp-agent trap enable feature-name） |
| Syslog設定 | 送信元インターフェース | 4 つとも | 対応（info-center loghost source） |
| 時刻設定 | NTP サーバーの優先指定（prefer） | 4 つとも | 対応（ntp unicast-server … preferred） |
| STP設定 | インスタンスごとの VLAN 割り当てと優先度 | NP3000・NP4000 | 対応（instance、stp instance priority） |
| STP設定 | STP を止める VLAN／ポート | C9300L・NP3000・NP4000 | 対応（ポートで stp disable） |
| STP設定 | ポートごとのコスト | 4 つとも | 対応 |
| ポート設定 | Voice VLAN | C9200L | 対応（voice-vlan） |
| ポート設定 | ループ検知 | NP3000 | 対応（loopback-detect）。**自動モードが既定で有効** |
| システム設定 | IGMP スヌーピング | C9300L | 対応（igmp snooping enable、既定は無効） |
| システム設定 | Web 管理（HTTP／HTTPS）の無効化 | C9200L・C9300L | 対応（web-manager）。**既定で HTTPS 8443 と HTTP から HTTPS への転送が有効** |
| ACL設定 | VLAN へのフィルター適用（vlan access-map／vlan filter） | NP3000 | 未調査 |
| ハードウェア構成情報 | ライセンス、冗長電源、オプションモジュールの行 | C9200L・C9300L・NP3000 | ライセンスは不要。冗長電源は未調査 |

## 3. ver0.7 で直したほうがよい点（今回わかったこと）

- **Routed Port設定**: Huawei の YunShan 系マニュアルでは、物理ポートを L3 に切り替えるコマンド（undo portswitch）が使えるのは S6780-H などに限られます。S5735-L-V2 は対象外と読めるため、節の削除を検討してください（実機での確認を推奨）。L3 は Vlanif で設定します。
- **システム初期設定の Web Manager**: V600 のコマンド名は web-manager enable です。設計値の欄を足すときは、この名前で書いてください。

## 4. 確認できなかった点

- RRPP（リング冗長の 1 方式）は、V600R025C00 のマニュアルに章もコマンドも見つかりませんでした。
- S5735-L-V2 がライセンス手引きの「S57XX-L Series」に当たるかは、表に系列名しかなく推測です。
- ip device tracking の相当機能、ACL の VLAN フィルター、冗長電源の対応は調べていません。

## 5. 根拠

Huawei 側はすべて CloudEngine S3700/S5700/S6700 V600R025C00 の公式マニュアルです。各機能の「Configuration Precautions」ページにある対応機種表で、S5735-L-V2 の行を確認しました。

| 機能 | URL |
|---|---|
| DNS | https://support.huawei.com/enterprise/en/doc/EDOC1100515490/b4b8e9fa/configuration-precautions-for-dns |
| ジャンボフレーム／MTU | https://support.huawei.com/enterprise/en/doc/EDOC1100515479/371b10df/configuring-general-attributes |
| AAA（RADIUS／HWTACACS） | https://support.huawei.com/enterprise/en/doc/EDOC1100515497/f747a035/configuration-precautions-for-aaa |
| DHCP | https://support.huawei.com/enterprise/en/doc/EDOC1100515490/a1f24f00/configuration-precautions-for-dhcpv4 |
| LoopBack | https://support.huawei.com/enterprise/en/doc/EDOC1100515479/99be6217/configuration-precautions-for-logical-interface |
| NetStream（非対応） | https://support.huawei.com/enterprise/en/doc/EDOC1100512791/e29a0bb0/configuration-precautions-for-netstream |
| sFlow | https://support.huawei.com/enterprise/en/doc/EDOC1100512791/82b9055e/configuration-precautions-for-sflow |
| route-policy | https://support.huawei.com/enterprise/en/doc/EDOC1100515486/ab20b5fb/configuration-precautions-for-routing-policy |
| ERPS | https://support.huawei.com/enterprise/en/doc/EDOC1100515500/f318bd15/erps-configuration |
| SEP | https://support.huawei.com/enterprise/en/doc/EDOC1100515500/e7771f6e/configuration-precautions-for-sep |
| QinQ・Voice VLAN | https://support.huawei.com/enterprise/en/doc/EDOC1100515500/31f1ee38/configuration-precautions-for-vlan |
| NAC（802.1X・MAC・Portal） | https://support.huawei.com/enterprise/en/doc/EDOC1100515497/8698213e/configuration-precautions-for-nac |
| IGMP スヌーピング | https://support.huawei.com/enterprise/en/doc/EDOC1100515484/1bea1bbe/configuration-precautions-for-layer-2-multicast |
| ループ検知 | https://support.huawei.com/enterprise/en/doc/EDOC1100515500/b15a9314/default-settings-for-loopback-detection |
| ライセンス | https://support.huawei.com/enterprise/en/doc/EDOC1100277353/a0fb1667/conventional-license-control-items |
| Web 管理 | https://support.huawei.com/enterprise/en/doc/EDOC1100515494/e6d4ce15/configuration-precautions-for-web-ui-based-login |
| Syslog 送信元 | https://support.huawei.com/enterprise/en/doc/EDOC1100515474/682c75a0/configuring-the-device-to-output-logs-to-a-log-host |
| NTP 優先指定 | https://support.huawei.com/enterprise/en/doc/EDOC1100515397/6191e055/ntp-configuration-commands |
| SNMPv2c・Trap | https://support.huawei.com/enterprise/en/doc/EDOC1100515474/5912a865/configuration-precautions-for-snmp |
| MSTP インスタンス | https://support.huawei.com/enterprise/en/doc/EDOC1100515500/8849da2/configuring-an-mst-region-for-mstp |
| portswitch（L3 切り替え） | https://info.support.huawei.com/hedex/api/pages/EDOC1100363264/AEN0403J/05/resources/command/yunshan/PORTSWITCH(VLANOM).html |

