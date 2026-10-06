"""パラメーターシートの値。機種ごとの値（MODELS）と全機種共通の値（COMMON）。

"要記入" で始まる値は、生成時に黄色で塗る（利用者の記入・実機確認が要る欄）。
"""

REVISIONS = [
    {
        "version": "ver0.7",
        "date": "2026/10/05",
        "kind": "改訂",
        "summary": "Huawei（SphereOS）の用語・既定値へ統一。\n"
                   "C48T4XE-L の崩れ、行22～24 の表示切れ、10GE ポート図の番号を修正。\n"
                   "スタック構成情報、メンバー2～4 のポート設定、SNMPv3、PoE 設定欄を追加。",
    },
    {
        "version": "ver0.8",
        "date": "2026/10/05",
        "kind": "改訂",
        "summary": "DNS設定、MTU設定、RADIUS設定、DHCP設定を追加。",
    },
]

SW_VERSION = "代表機確認値：SphereOS 1.25.0.1 / V600R025C00SPC500"
SW_FILE = "代表機確認値：S5735-V2_V600R025C00SPC500.cc"

COMMON = {
    # スタック
    "stack_priority": "100（default）",
    "stack_domain": "未設定（default）",
    "stack_dad": "要記入（direct／relay）",
    "stack_note": "インターフェース名の先頭の数字がメンバーID（1～9、既定 1）。例: メンバー2 の 1 番ポートは GE2/0/1。\n"
                  "優先度は 1～255（既定 100）。値が大きい筐体ほどマスターに選ばれやすい。\n"
                  "最大 9 台まで組める。5 台以上なら列を足す。\n"
                  "スタックの設定は running-config に出ない（display stack configuration で確認する）。",
    # 管理アクセス
    "console_auth": "password（実機確認値）",
    "console_ui": "user-interface con 0",
    "vty": "user-interface vty 0 4（5 本）",
    "vty_auth": "aaa／SSH",
    "stelnet": "無効（default）",
    "ssh_source": "要記入（管理用Vlanif 等）",
    "rsa_min": "3072",
    "ssh_ver": "2.0",
    "ssh_timeout": "60秒",
    "ssh_retry": "3回",
    "acl_ssh": "SSH：未設定",
    "acl_web": "Web：未設定",
    "acl_snmp": "SNMP：未設定",
    "mgmt_note": "コンソールは初回ログイン時にパスワード設定が必須。\n"
                 "SSH を使うには stelnet server enable と ssh server-source の指定が要る（既定ではどのインターフェースからも受け付けない）。\n"
                 "VTY の認証方式は aaa にする。既定で使える VTY は 0～4（最大 21 本）。\n"
                 "RSA 鍵は rsa local-key-pair create で作る（既定 3072 ビット）。",
    # 時刻・Syslog
    "timezone": "UTC（default）",
    "ntp": "未設定",
    "time_note": "国内に設置する場合は clock timezone JST add 09:00:00 を設定する。",
    "syslog": "未設定",
    "syslog_level": None,
    "syslog_facility": None,
    "syslog_note": None,
    # DNS
    "dns_resolve": "無効（default）",
    "dns_note": "DNS サーバーは最大 6 個（IPv4・IPv6 の合計）。",
    # MTU
    "jumbo": "9216 バイト（default）",
    "vlanif_mtu": "1500 バイト（default）",
    "mtu_note": "最大フレーム長は 1518～10240 バイト、MTU は 46～9600 バイトの範囲で設定する。\n"
                "MTU は Vlanif ごとに設定する。",
    # RADIUS
    "radius_note": "HWTACACS を使う場合も同じ欄に記入する（hwtacacs-server template）。\n"
                   "認証の順序の例: radius local",
    # DHCP
    "dhcp_enable": "無効（default）",
    "dhcp_snooping": "無効（default）",
    "dhcp_note": "リレーは Vlanif で dhcp select relay と dhcp relay server-ip を設定する。\n"
                 "サーバーは ip pool でアドレスプールを作る。",
    # SNMP
    "snmp_agent": "無効（default）",
    "snmp_version": "v3",
    "snmp_auth": None,
    "snmp_priv": None,
    "snmp_level": "privacy（認証＋暗号）",
    "snmp_note": "認証方式は sha2-256／sha2-384／sha2-512、暗号方式は aes128／aes192／aes256 から選ぶ。\n"
                 "md5・sha・sha2-224・des56・3des168 は WEAKEA（弱い暗号方式の追加パッケージ）が要る。\n"
                 "Trap を送るには snmp-agent trap enable も設定する。",
    # STP
    "stp_mode": "MSTP（default）",
    "stp_priority": "32768（default）",
    "stp_hello": "2秒／20秒（default）",
    "stp_fwd": "15秒／20（default）",
    "bpdu_protection": "無効（default）",
    "region_name": "管理IFのMACアドレス（default）",
    "region_rev": "0（default）",
    "region_map": "MSTI 0：VLAN 1～4094（default）",
    "port_priority": "128（default）",
    "path_cost": "自動計算（default）",
    "edged_default": "無効（default）",
    "stp_note": "BPDU 保護（stp bpdu-protection）はシステム全体の設定。\n"
                "エッジポート（stp edged-port enable）・BPDU Filter・ループ保護はポートごとに設定する。\n"
                "ループ保護とルート保護は同じポートに併用できない。",
    # LLDP
    "lldp": "有効（default）",
    "lldp_note": None,
    # 管理用インターフェース
    "mgmt_if_note": "専用の物理管理ポートではない。管理IPは Vlanif で設定する（管理用Vlanif の欄に記入）。",
    "system_defaults": [
        "Information Center：有効",
        "Log Buffer：有効（実サイズ512／最大10240）",
        "Trap Buffer：有効（実サイズ256／最大1024）",
        "Web Manager：有効、HTTPSポート8443",
        "Telnet Server Source：全インターフェース無効",
        "SSH Server Source：全インターフェース無効",
    ],
    "vlan_note": "VLAN 1 が既定の VLAN。",
    "vrrp_note": "VRRP の既定値：優先度 100、通知間隔 1 秒、プリエンプト遅延 0 秒（起動時・インターフェース Up 時は 5 秒）、"
                 "監視インターフェースのダウン時の優先度減算 10。既定値のままなら空欄でよい。",
    # ポート
    "speed_ge": "Auto",
    "duplex_ge": "Auto",
    "speed_x": "10G",
    "duplex_x": "Full",
    "link_type": "negotiation-auto",
    "allow_vlan": "-",
    "updown_trap": "enable",
    "port_note": "link-type の既定は negotiation-auto（対向との交渉で access か trunk かを自動で決める）。固定する場合は access／trunk／hybrid を記入する。\n"
                 "GE ポートは自動ネゴシエーションが既定で有効（Speed・Duplex の Auto）。10GE ポートに 10GE 光モジュールを挿した場合は自動ネゴシエーション非対応。",
    # PoE
    "poe_enable": "有効（default）",
    "poe_priority": "low（default）",
    "poe_power": "30000 mW（default）",
    "poe_note": "給電優先度は critical／high／low。ポート上限電力は mW 単位（0～90000）。",
    # Eth-Trunk
    "trunk_lb": "IPパケット：送信元/宛先IP・L4送信元/宛先ポート、L2パケット：送信元/宛先MAC（default）",
    "trunk_note": "初期状態では Eth-Trunk は未設定。動作モードの既定は manual（LACP を使う場合は lacp-static）。\n"
                  "負荷分散方式はシステム全体の load-balance profile で決める（Eth-Trunk ごとには設定できない）。",
    # ルーティング
    "static_pref": "60",
    "routing_note": "OSPF の network はワイルドカードマスクで指定する（例: network 192.168.1.0 0.0.0.255）。",
    # ACL・QoS・ストーム制御
    "acl_note": "ACL番号は 基本ACL 2000～2999、拡張ACL 3000～3999。名前付き ACL も使える。",
    "qos_note": "MQC（流分類・流動作・流ポリシー）で設定し、インターフェースへ適用する。\n"
                "trust は 8021p（inner／outer）か dscp。キュースケジューリングは qos schedule-profile で設定する（既定 PQ。PQ と DRR を使える）。",
    "storm_recovery": "未設定",
    "storm_note": "action は error-down／block／suppress。interval は 1～180 秒（既定 5 秒）。log・trap は既定で無効。\n"
                  "受信方向のトラフィック抑制が既定で有効（ブロードキャスト 10%）。同じ種類のパケットにはストーム制御と併用できない。",
}

MODELS = [
    {
        "sheet": "C8T4X-Q-L",
        "ge": 8, "xge": 4,
        "port_summary": "10/100/1000BASE-T × 8、10GE SFP+ × 4",
        "sw_version": SW_VERSION, "sw_file": SW_FILE,
        "power": "内蔵AC電源", "cooling": "ファンレス",
        "poe_std": "非対応", "poe_ports": "－", "poe_budget": "非対応",
        "stack_port_hw": "専用スタックポートなし（業務ポートを使う）",
        "stack_port_cfg": "要記入",
        "dedicated_stack": False,
        "hw_note": "ハードウェア差分は製品データシートを反映。個体値・型番表示・OSファイルは実機未確認。\n"
                   "共通の初期値は C8P4X-Q-L 実機（SphereOS 1.25.0.1 / V600R025C00SPC500）の確認結果を横展開。",
    },
    {
        "sheet": "C8P4X-Q-L",
        "ge": 8, "xge": 4,
        "port_summary": "10/100/1000BASE-T × 8（PoE）、10GE SFP+ × 4",
        "sw_version": "SphereOS 1.25.0.1 / V600R025C00SPC500",
        "sw_file": SW_FILE,
        "power": "要記入", "cooling": "ファンレス",
        "poe_std": "要記入", "poe_ports": "GE1/0/1～8", "poe_budget": "要記入",
        "stack_port_hw": "要記入",
        "stack_port_cfg": "要記入",
        "dedicated_stack": False,
        "hw_note": "実機確認済み：C8P4X-Q-L / SphereOS 1.25.0.1 / V600R025C00SPC500。",
    },
    {
        "sheet": "C24T4XE-L",
        "ge": 24, "xge": 4,
        "port_summary": "10/100/1000BASE-T × 24、10GE SFP+ × 4",
        "sw_version": "SphereOS 1.25.0.1 / V600R025C00SPC500",
        "sw_file": SW_FILE,
        "power": "要記入（実機表示：PWR1 正常）", "cooling": "要記入（実機表示：FAN1 正常）",
        "poe_std": "非対応", "poe_ports": "－", "poe_budget": "非対応",
        "stack_port_hw": "専用スタックポート × 2",
        "stack_port_cfg": "専用（設定不要）",
        "dedicated_stack": True,
        "hw_note": "実機確認済み：C24T4XE-L / SphereOS 1.25.0.1 / V600R025C00SPC500。専用スタックポートのCLI表示も確認済み。\n"
                   "実機表示：Memory 2048MB / Flash 1024MB。",
        "diagram_note": "STACK は専用スタックポート。",
    },
    {
        "sheet": "C24P4XE-L",
        "ge": 24, "xge": 4,
        "port_summary": "10/100/1000BASE-T × 24（PoE+）、10GE SFP+ × 4",
        "sw_version": SW_VERSION, "sw_file": SW_FILE,
        "power": "内蔵AC電源", "cooling": "ファン × 2",
        "poe_std": "PoE+", "poe_ports": "GE1/0/1～24", "poe_budget": "総給電能力 400W",
        "stack_port_hw": "専用スタックポート（12GE）× 2",
        "stack_port_cfg": "専用（設定不要）",
        "dedicated_stack": True,
        "hw_note": "ハードウェア差分は製品データシートを反映。スタックCLI表示は C24T4XE-L 代表機の結果を参考記載。",
        "diagram_note": "STACK は専用スタックポート。",
    },
    {
        "sheet": "C48T4XE-L",
        "ge": 48, "xge": 4,
        "port_summary": "10/100/1000BASE-T × 48、10GE SFP+ × 4",
        "sw_version": SW_VERSION, "sw_file": SW_FILE,
        "power": "内蔵AC電源", "cooling": "ファン × 1",
        "poe_std": "非対応", "poe_ports": "－", "poe_budget": "非対応",
        "stack_port_hw": "専用スタックポート（12GE）× 2",
        "stack_port_cfg": "専用（設定不要）",
        "dedicated_stack": True,
        "hw_note": "ハードウェア差分は製品データシートを反映。C24T4XE-L でスタック動作表示を代表確認済み。C48 個体の CLI 名は未確認。\n"
                   "専用スタックポート 2 ポートはデータシートで確認済み。CLI 上の物理インターフェース名は当該機種の実機では未確認。",
        "diagram_note": "STACK は専用スタックポート（CLI 名は未確認）。",
    },
    {
        "sheet": "C48LP4XE-L",
        "ge": 48, "xge": 4,
        "port_summary": "10/100/1000BASE-T × 48（PoE+）、10GE SFP+ × 4",
        "sw_version": SW_VERSION, "sw_file": SW_FILE,
        "power": "内蔵AC電源", "cooling": "ファン × 2",
        "poe_std": "PoE+", "poe_ports": "GE1/0/1～48", "poe_budget": "総給電能力 380W",
        "stack_port_hw": "専用スタックポート（12GE）× 2",
        "stack_port_cfg": "専用（設定不要）",
        "dedicated_stack": True,
        "hw_note": "ハードウェア差分は製品データシートを反映。",
        "diagram_note": "STACK は専用スタックポート（CLI 名は未確認）。",
    },
]
