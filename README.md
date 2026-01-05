# Excel 智能助手 Add-in

涓€涓己澶х殑Excel鍔犺浇椤癸紝闆嗘垚浜咥I鍔熻兘锛屾彁渚涙櫤鑳芥暟鎹垎鏋愬拰鑷姩鍖栨搷浣溿€?

## 鍔熻兘鐗规€?

### 馃 AI鏅鸿兘鍔╂墜
- **鑷劧璇█澶勭悊**锛氫娇鐢ㄨ嚜鐒惰瑷€鎻忚堪Excel鎿嶄綔锛岃嚜鍔ㄦ墽琛?
- **鏁版嵁娲炲療**锛欰I鍒嗘瀽鏁版嵁妯″紡骞舵彁渚涘缓璁?
- **鏅鸿兘鏍煎紡鍖?*锛氳嚜鍔ㄨ瘑鍒暟鎹被鍨嬪苟搴旂敤鏈€浣虫牸寮?

### 馃洜锔?鍛戒护鍔熻兘
- **鏁版嵁娓呯悊涓庢牸寮忓寲**锛氫竴閿竻鐞嗗拰鏍煎紡鍖栭€夊畾鏁版嵁
- **AI娲炲療鐢熸垚**锛氫负閫夊畾鏁版嵁鐢熸垚鏅鸿兘鍒嗘瀽鎶ュ憡
- **琛ㄦ牸杞崲**锛氬皢閫夊畾鑼冨洿杞崲涓轰笓涓氭牸寮忕殑Excel琛ㄦ牸
- **宸ヤ綔绨挎憳瑕?*锛氬垱寤哄寘鍚叧閿寚鏍囩殑宸ヤ綔绨挎憳瑕佽〃

### 馃帹 鐢ㄦ埛鐣岄潰
- **鐜颁唬鍖栬璁?*锛氫娇鐢‵luent UI缁勪欢搴?
- **涓婚鏀寔**锛氭敮鎸佹祬鑹层€佹殫鑹插拰楂樺姣斿害涓婚
- **鍝嶅簲寮忓竷灞€**锛氶€傞厤涓嶅悓灞忓箷灏哄
- **IE鍏煎**锛氭敮鎸両E 10+娴忚鍣?

## 鎶€鏈爤

- **鍓嶇**锛歊eact 19 + TypeScript
- **UI缁勪欢**锛欯fluentui/react-components
- **鏋勫缓宸ュ叿**锛歐ebpack 5 + Babel
- **鍚庣**锛欵xpress.js + DeepSeek AI API
- **Office闆嗘垚**锛歄ffice JavaScript API

## 椤圭洰缁撴瀯

```
excel-copilot-addin/
鈹溾攢鈹€ src/
鈹?  鈹溾攢鈹€ taskpane/
鈹?  鈹?  鈹溾攢鈹€ components/     # React缁勪欢
鈹?  鈹?  鈹?  鈹溾攢鈹€ App.tsx     # 涓诲簲鐢ㄧ粍浠?
鈹?  鈹?  鈹?  鈹溾攢鈹€ Header.tsx  # 澶撮儴缁勪欢
鈹?  鈹?  鈹?  鈹溾攢鈹€ HeroList.tsx # 鍔熻兘鍒楄〃
鈹?  鈹?  鈹?  鈹斺攢鈹€ Progress.tsx # 杩涘害鎸囩ず鍣?
鈹?  鈹?  鈹溾攢鈹€ taskpane.ts     # 浠诲姟绐楁牸鍏ュ彛
鈹?  鈹?  鈹溾攢鈹€ taskpane.html   # HTML妯℃澘
鈹?  鈹?  鈹斺攢鈹€ taskpane.css    # 鏍峰紡鏂囦欢
鈹?  鈹溾攢鈹€ commands/
鈹?  鈹?  鈹溾攢鈹€ commands.ts     # Office鍛戒护鍔熻兘
鈹?  鈹?  鈹斺攢鈹€ commands.html   # 鍛戒护椤甸潰
鈹?  鈹斺攢鈹€ ...
鈹溾攢鈹€ ai-backend.js           # AI鍚庣鏈嶅姟
鈹溾攢鈹€ mock-backend/           # 妯℃嫙鍚庣
鈹溾攢鈹€ webpack.config.js       # Webpack閰嶇疆
鈹溾攢鈹€ package.json           # 椤圭洰閰嶇疆
鈹溾攢鈹€ manifest.xml           # Office鍔犺浇椤规竻鍗?
鈹斺攢鈹€ README.md              # 椤圭洰鏂囨。
```

## 蹇€熷紑濮?

### 鍓嶆彁鏉′欢
- Node.js 16+
- Excel 2016+ 鎴?Excel Online
- DeepSeek API瀵嗛挜锛堝彲閫夛紝鐢ㄤ簬AI鍔熻兘锛?

### 瀹夎姝ラ

1. **鍏嬮殕椤圭洰**
   ```bash
   git clone <repository-url>
   cd excel-copilot-addin
   ```

2. **瀹夎渚濊禆**
   ```bash
   npm install
   ```

3. **閰嶇疆鐜鍙橀噺**
   - 澶嶅埗鐜鍙橀噺妯℃澘锛?
     ```bash
     cp .env.example .env
     ```
   - 缂栬緫 `.env` 鏂囦欢锛屽～鍐欐偍鐨凞eepSeek API瀵嗛挜锛?
     ```
     DEEPSEEK_API_KEY=your_deepseek_api_key_here
     DEEPSEEK_API_BASE=https://api.deepseek.com
     DEEPSEEK_MODEL=deepseek-chat
     ```
   - 濡傛灉娌℃湁API瀵嗛挜锛屽彲浠ユ殏鏃朵娇鐢ㄦā鎷熸暟鎹ā寮?

4. **鍚姩寮€鍙戞湇鍔″櫒**
   ```bash
   npm run dev:full
   ```
   杩欏皢鍚屾椂鍚姩锛?
   - 寮€鍙戞湇鍔″櫒 (绔彛3000)
   - AI鍚庣鏈嶅姟 (绔彛3001)
   - 妯℃嫙鍚庣 (绔彛3002)

5. **鏃佸姞杞藉姞杞介」**
   - 鎵撳紑Excel
   - 杞埌"寮€鍙戝伐鍏?閫夐」鍗?
   - 鐐瑰嚮"鏃佸姞杞藉姞杞介」"
   - 閫夋嫨 `manifest.xml` 鏂囦欢

## 鍙敤鑴氭湰

```bash
# 寮€鍙戞ā寮忥紙瀹屾暣锛?
npm run dev:full

# 浠呭紑鍙戞湇鍔″櫒
npm run dev-server

# 浠匒I鍚庣
npm run ai-backend

# 鏋勫缓鐢熶骇鐗堟湰
npm run build

# 鏋勫缓骞跺垎鏋愬寘澶у皬
npm run build:analyze

# 杩愯娴嬭瘯
npm run test

# 杩愯娴嬭瘯骞剁敓鎴愯鐩栫巼鎶ュ憡
npm run test:coverage

# TypeScript绫诲瀷妫€鏌?
npm run type-check

# 浠ｇ爜鏍煎紡鍖栧拰lint妫€鏌?
npm run format

# 瀹夊叏妫€鏌?
npm run security-check
```

## 閰嶇疆璇存槑

### AI API閰嶇疆
椤圭洰鏀寔DeepSeek AI API锛屾偍闇€瑕侊細
1. 娉ㄥ唽DeepSeek璐︽埛骞惰幏鍙朅PI瀵嗛挜
2. 鍦ㄥ簲鐢ㄤ腑閰嶇疆API瀵嗛挜
3. 鎴栦娇鐢ㄥ唴缃殑妯℃嫙鏁版嵁妯″紡

### 涓婚閰嶇疆
鏀寔涓夌涓婚妯″紡锛?
- **娴呰壊涓婚**锛氶粯璁や富棰?
- **鏆楄壊涓婚**锛氶€傚悎澶滈棿浣跨敤
- **楂樺姣斿害涓婚**锛氭彁楂樺彲璁块棶鎬?

### 娴忚鍣ㄥ吋瀹规€?
- Chrome 80+
- Firefox 75+
- Edge 80+
- Safari 13+
- IE 10+锛堟湁闄愭敮鎸侊級

## 寮€鍙戞寚鍗?

### 娣诲姞鏂板姛鑳?
1. 鍦?`src/taskpane/components/` 涓垱寤篟eact缁勪欢
2. 鍦?`src/commands/commands.ts` 涓坊鍔燨ffice鍛戒护
3. 鏇存柊 `manifest.xml` 娉ㄥ唽鏂板懡浠?
4. 娣诲姞鐩稿簲鐨勬牱寮忓埌 `src/taskpane/taskpane.css`

### 璋冭瘯鎶€宸?
- 浣跨敤 `npm run dev:full` 鍚姩瀹屾暣寮€鍙戠幆澧?
- 鍦‥xcel涓寜F12鎵撳紑寮€鍙戣€呭伐鍏?
- 鏌ョ湅鎺у埗鍙版棩蹇椾簡瑙PI璋冪敤鎯呭喌
- 浣跨敤React Developer Tools璋冭瘯缁勪欢

## 鏁呴殰鎺掗櫎

### 甯歌闂

1. **鍔犺浇椤规棤娉曞姞杞?*
   - 妫€鏌anifest.xml璺緞鏄惁姝ｇ‘
   - 纭寮€鍙戞湇鍔″櫒姝ｅ湪杩愯
   - 妫€鏌ユ祻瑙堝櫒鎺у埗鍙伴敊璇?

2. **AI鍔熻兘涓嶅彲鐢?*
   - 纭API瀵嗛挜宸查厤缃?
   - 妫€鏌ョ綉缁滆繛鎺?
   - 鏌ョ湅鍚庣鏈嶅姟鏃ュ織

3. **鏍峰紡闂**
   - 娓呴櫎娴忚鍣ㄧ紦瀛?
   - 妫€鏌SS鏂囦欢鏄惁姝ｇ‘鍔犺浇
   - 楠岃瘉涓婚绫绘槸鍚︽纭簲鐢?

### 鏃ュ織鏌ョ湅
- 鍓嶇鏃ュ織锛氭祻瑙堝櫒鎺у埗鍙?
- 鍚庣鏃ュ織锛氱粓绔緭鍑?
- AI鏈嶅姟鏃ュ織锛氱鍙?001鐨勬帶鍒跺彴

## 璐＄尞鎸囧崡

1. Fork椤圭洰
2. 鍒涘缓鍔熻兘鍒嗘敮 (`git checkout -b feature/AmazingFeature`)
3. 鎻愪氦鏇存敼 (`git commit -m 'Add some AmazingFeature'`)
4. 鎺ㄩ€佸埌鍒嗘敮 (`git push origin feature/AmazingFeature`)
5. 鎵撳紑Pull Request

## 璁稿彲璇?

鏈」鐩熀浜嶮IT璁稿彲璇?- 鏌ョ湅 [LICENSE](LICENSE) 鏂囦欢浜嗚В璇︽儏

## 鏀寔

- 闂鎶ュ憡锛歔GitHub Issues](https://github.com/yourusername/excel-copilot-addin/issues)
- 鍔熻兘璇锋眰锛歔GitHub Discussions](https://github.com/yourusername/excel-copilot-addin/discussions)
- 鏂囨。锛歔椤圭洰Wiki](https://github.com/yourusername/excel-copilot-addin/wiki)

## 鏇存柊鏃ュ織

### v1.0.0 (2025-12-28)
- 鍒濆鐗堟湰鍙戝竷
- 鍩虹AI闆嗘垚鍔熻兘
- 瀹屾暣鐨凟xcel鎿嶄綔鍛戒护
- 澶氫富棰樻敮鎸?
- 鍝嶅簲寮忚璁?

---

**娉ㄦ剰**锛氳繖鏄竴涓紑鍙戜腑鐨勯」鐩紝鍔熻兘鍙兘浼氬彂鐢熷彉鍖栥€傚缓璁畾鏈熸洿鏂板埌鏈€鏂扮増鏈€?


