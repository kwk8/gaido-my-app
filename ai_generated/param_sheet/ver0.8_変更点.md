# LinkSphere パラメーターシート ver0.8 変更点

- 新ファイル: `LinkSphere_パラメーターシート_ver0.8_Huawei準拠.xlsx`
- 元: ver0.7（変更点は `ver0.7_変更点.md`）
- 足した理由: 社内の Cisco・Apresia の設定シートにあり、ver0.7 に無かった大項目のうち、利用者が採用を決めた 4 つ（比較は `設定シート比較_Cisco_Apresia.md`）

## 足した節

全 6 機種のシートに、同じ内容で足しました。並び順は Cisco Catalyst 9300L の設定シートに合わせています。

| 節 | 置き場所 | 欄 | 既定値（Huawei V600R025C00） |
|---|---|---|---|
| DNS設定 | 時刻設定の後 | DNS 名前解決（dns resolve）、送信元 IP、DNS サーバー 1・2 | 名前解決は無効。サーバーは最大 6 個 |
| MTU設定 | Syslog設定の後 | 最大フレーム長（jumboframe enable）、Vlanif の MTU、MTU を変える Vlanif（2 行） | 最大フレーム長 9216 バイト、MTU 1500 バイト |
| RADIUS設定 | SNMP設定の後 | サーバーテンプレート名、認証サーバー 1・2、アカウンティングサーバー 1・2（IP・Port・共有鍵）、認証の順序、アカウンティング、適用するユーザー（ドメイン） | 未設定。HWTACACS も同じ欄で使える |
| DHCP設定 | Vlanif設定・Routed Port設定の後 | DHCP 機能（dhcp enable）、DHCP リレー（3 行）、DHCP サーバーのアドレスプール（2 行）、DHCP スヌーピング・信頼ポート | すべて無効 |

## 根拠

- DNS: https://support.huawei.com/enterprise/en/doc/EDOC1100515490/fc873161/default-settings-for-dns
- MTU（最大フレーム長）: https://support.huawei.com/enterprise/en/doc/EDOC1100515479/371b10df/configuring-general-attributes
- MTU（インターフェース）: https://support.huawei.com/enterprise/en/doc/EDOC1100515479/1a1ac02/configuring-an-mtu-for-an-interface
- RADIUS: https://support.huawei.com/enterprise/en/doc/EDOC1100515497/f8c774b8/configuring-a-radius-server-template
- HWTACACS: https://support.huawei.com/enterprise/en/doc/EDOC1100515497/59888353/configuring-an-hwtacacs-server-template
- DHCP リレー: https://support.huawei.com/enterprise/en/doc/EDOC1100515490/e2e45b87/configuring-dhcpv4-relay
- DHCP サーバー: https://support.huawei.com/enterprise/en/doc/EDOC1100515490/c9c90844/configuring-the-dhcpv4-server-function
- DHCP スヌーピング: https://support.huawei.com/enterprise/en/doc/EDOC1100515490/2fc8204a/configuration-precautions-for-dhcp-snooping

## 既定値を書かなかった欄

RADIUS の Port は、既定のポート番号をマニュアルで確認していないため空欄にしました。
