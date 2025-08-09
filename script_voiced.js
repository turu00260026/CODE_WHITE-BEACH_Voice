// ===== グローバル変数 (データと状態のみ) =====
let gameState = {};
let scenarioData = null;
let elements = {}; // 初期化時に設定するため、ここでは空
let typeInterval;
let currentAudio = null;

// ===== ★★★ ゲーム全体の実行開始点 ★★★ =====
// HTMLドキュメントの読み込みと解析が完了したら、すべての処理を開始する
document.addEventListener('DOMContentLoaded', async () => {
    // 1. DOM要素をここで初めて取得する（最重要）
    elements = {
        titleScreen: document.getElementById('title-screen'),
        protagonistSelect: document.getElementById('protagonist-select'),
        gameScreen: document.getElementById('game-screen'),
        menuScreen: document.getElementById('menu-screen'),
        endingScreen: document.getElementById('ending-screen'),
        loadingScreen: document.getElementById('loading-screen'),
        backgroundImage: document.getElementById('background-image'),
        characterImage: document.getElementById('character-image'),
        textbox: document.getElementById('textbox'),
        speakerNameContainer: document.getElementById('speaker-name-container'),
        speakerName: document.getElementById('speaker-name'),
        dialogueText: document.getElementById('dialogue-text'),
        nextIndicator: document.getElementById('next-indicator'),
        choicesContainer: document.getElementById('choices-container'),
        startButton: document.getElementById('start-button'),
        hiddenScenarioButton: document.getElementById('hidden-scenario-button'),
        shouButton: document.getElementById('shou-button'),
        syouButton: document.getElementById('syou-button'),
        menuButton: document.getElementById('menu-button'),
        backButton: document.getElementById('back-button'),
        resumeButton: document.getElementById('resume-button'),
        titleReturnButton: document.getElementById('title-return-button'),
        endingTitleButton: document.getElementById('ending-title-button'),
        endingTitle: document.getElementById('ending-title'),
        endingText: document.getElementById('ending-text')
    };

    // 2. シナリオデータを非同期で読み込む
    await loadScenarioData();
    
    // 3. 全てのイベントリスナーを設定する
    setupEventListeners();

    // 4. 初期画面（タイトル）を表示する
    showScreen(elements.titleScreen);
    checkHiddenScenarioUnlock();
    console.log('ゲームエンジン初期化完了');
});

// ===== データ読み込み =====
async function loadScenarioData() {
    try {
        showScreen(elements.loadingScreen);
        const response = await fetch('scenario_voiced.json');
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
        scenarioData = await response.json();
        console.log('シナリオデータ読み込み完了');
    } catch (error) {
        console.error('シナリオデータの読み込みに失敗:', error);
        alert('ゲームデータの読み込みに失敗しました。');
    }
}

// ===== イベントリスナー設定 =====
function setupEventListeners() {
    elements.startButton.addEventListener('click', () => showScreen(elements.protagonistSelect));
    elements.hiddenScenarioButton.addEventListener('click', () => selectProtagonist('隠しシナリオ'));
    elements.shouButton.addEventListener('click', () => selectProtagonist('ショウ編'));
    elements.syouButton.addEventListener('click', () => selectProtagonist('しょう編'));
    elements.gameScreen.addEventListener('click', handleGameScreenClick);
    elements.menuButton.addEventListener('click', () => elements.menuScreen.classList.remove('hidden'));
    elements.backButton.addEventListener('click', goBack);
    elements.resumeButton.addEventListener('click', () => elements.menuScreen.classList.add('hidden'));
    elements.titleReturnButton.addEventListener('click', returnToTitle);
    elements.endingTitleButton.addEventListener('click', returnToTitle);
}

// ===== 画面制御 =====
function showScreen(screenElement) {
    document.querySelectorAll('.screen').forEach(s => s.classList.add('hidden'));
    if (screenElement) screenElement.classList.remove('hidden');
}

// ===== ゲーム開始処理 =====
function selectProtagonist(protagonistKey) {
    gameState = {
        currentProtagonist: protagonistKey,
        currentBlock: protagonistKey === '隠しシナリオ' ? 'main' : 'act1',
        currentIndex: 0,
        totalScore: 0,
        lastChoiceScore: 0,
        isPlaying: true,
        isChoiceScene: false,
        currentCharacterImage: '',
        currentBackground: '',
        history: []
    };
    showScreen(elements.gameScreen);
    advanceScene();
}

// ===== シナリオ進行の心臓部 =====
function advanceScene() {
    // 前の音声を停止
    stopCurrentVoice();

    if (gameState.isChoiceScene) return;
    
    // scenarioData[protagonist][block][index] のパスで行データを取得
    const currentBlockData = scenarioData[gameState.currentProtagonist][gameState.currentBlock];
    
    if (!currentBlockData || gameState.currentIndex >= currentBlockData.length) {
        // ブロックの終端に達した場合
        const lastLine = currentBlockData ? currentBlockData[currentBlockData.length - 1] : {};
        
        // 次のブロックへの遷移を試す
        const hasTransition = handleBlockEnd(lastLine);
        
        // 遷移がない場合は、エンディングブロックかどうかチェック
        if (!hasTransition) {
            // エンディングブロックまたは進行可能な次のブロックがない場合はエンディング表示
            if (isEndingBlock(gameState.currentBlock) || lastLine.ending || lastLine.type === 'ending') {
                showEnding();
            } else {
                // デフォルトでエンディング処理
                showEnding();
            }
        }
        return;
    }
    
    const lineData = currentBlockData[gameState.currentIndex];
    displayLine(lineData);
    gameState.currentIndex++;
}

// エンディングブロックかどうかを判定
function isEndingBlock(blockName) {
    return blockName && (blockName.includes('ending_') || blockName === 'ending');
}

// ===== 画面表示の心臓部 =====
function displayLine(lineData) {
    if (lineData.type === 'choice') {
        displayChoice(lineData.options);
        return;
    }
    
    elements.choicesContainer.classList.add('hidden');
    elements.textbox.style.display = 'flex';
    
    // 背景画像の処理
    if (lineData.background) {
        gameState.currentBackground = lineData.background;
    }
    if (gameState.currentBackground) {
        elements.backgroundImage.src = gameState.currentBackground;
    }
    
    // キャラクター画像の処理
    if (lineData.char_image !== undefined) {
        gameState.currentCharacterImage = lineData.char_image;
    }
    elements.characterImage.src = gameState.currentCharacterImage || '';
    elements.characterImage.style.display = gameState.currentCharacterImage ? 'block' : 'none';
    
    // 話者名の処理
    const speaker = lineData.character || '';
    if (speaker && !speaker.includes('ナレーション')) {
        elements.speakerName.textContent = speaker;
        elements.speakerNameContainer.style.display = 'block';
    } else {
        elements.speakerNameContainer.style.display = 'none';
    }
    
    // 音声再生（voiceプロパティがある場合のみ）
    if (lineData.voice) {
        console.log(`音声再生: ${lineData.voice}`);
        playVoice(lineData.voice);
    }
    
    // テキスト表示
    displayTextWithTypewriter(lineData.text || '');
    updateBackButtonVisibility();
}

// ===== ブロック終端の処理（すべての遷移タイプ対応） =====
function handleBlockEnd(lastLine = {}) {
    // next_block による次のブロックへの遷移（最優先）
    if (lastLine.next_block) {
        gameState.currentBlock = lastLine.next_block;
        gameState.currentIndex = 0;
        advanceScene();
        return true; // 遷移が発生した
    }
    
    // next_act による次のブロックへの遷移
    if (lastLine.next_act) {
        gameState.currentBlock = lastLine.next_act;
        gameState.currentIndex = 0;
        advanceScene();
        return true; // 遷移が発生した
    }
    
    // branch_act3 による分岐処理
    if (lastLine.branch_act3) {
        let route;
        if (gameState.totalScore >= 4) {
            route = 'act3_high';
        } else if (gameState.totalScore >= 1) {
            route = 'act3_mid';
        } else {
            route = 'act3_low';
        }
        gameState.currentBlock = route;
        gameState.currentIndex = 0;
        advanceScene();
        return true; // 遷移が発生した
    }
    
    // branch_ending による最終エンディング分岐処理
    if (lastLine.branch_ending) {
        // 履歴に現在の位置を保存してから遷移
        addToHistory();
        
        let endingRoute;
        if (gameState.totalScore >= 6) {
            endingRoute = 'ending_true';
        } else if (gameState.totalScore >= 1) {
            // ノーマルエンドの場合、最後の選択肢の点数で分岐
            if (gameState.lastChoiceScore === 2) {
                endingRoute = 'ending_normal_from_true';
            } else if (gameState.lastChoiceScore === 0) {
                endingRoute = 'ending_normal_from_bad';
            } else {
                endingRoute = 'ending_normal';
            }
        } else {
            endingRoute = 'ending_bad';
        }
        gameState.currentBlock = endingRoute;
        gameState.currentIndex = 0;
        advanceScene();
        return true; // 遷移が発生した
    }
    
    return false; // 遷移が発生しなかった
}

// ===== 選択肢の処理 =====
function displayChoice(options) {
    gameState.isChoiceScene = true;
    elements.speakerNameContainer.style.display = 'none';
    elements.dialogueText.textContent = '';
    elements.nextIndicator.style.display = 'none';
    elements.choicesContainer.innerHTML = '';
    
    options.forEach(option => {
        const button = document.createElement('button');
        button.className = 'choice-button';
        button.textContent = option.text;
        button.onclick = () => selectChoice(option);
        elements.choicesContainer.appendChild(button);
    });
    
    elements.choicesContainer.classList.remove('hidden');
}

function selectChoice(option) {
    // 1. スコアを加算
    gameState.totalScore += option.score !== undefined ? option.score : 0;
    // 最後の選択肢のポイントを記録
    gameState.lastChoiceScore = option.score !== undefined ? option.score : 0;
    
    // 2. 選択肢モードを終了
    gameState.isChoiceScene = false;
    elements.choicesContainer.classList.add('hidden');
    elements.textbox.style.display = 'flex';
    
    // 3. 次のブロックへの遷移処理（必須）
    if (option.next_block) {
        // 履歴に現在の位置を保存
        addToHistory();
        
        // 新しいブロック（エンディングブロックを含む）に移動
        gameState.currentBlock = option.next_block;
        gameState.currentIndex = 0;
        
        // 新しいブロックの最初の行を表示
        advanceScene();
    } else {
        // next_blockがない場合は現在のシナリオを続行
        addToHistory();
        advanceScene();
    }
}

// ===== 履歴管理システム =====
function addToHistory() {
    // 位置情報とスコア情報を履歴に保存
    const positionData = {
        block: gameState.currentBlock,
        index: gameState.currentIndex,
        totalScore: gameState.totalScore,
        lastChoiceScore: gameState.lastChoiceScore
    };
    gameState.history.push(positionData);
    
    // 履歴が50を超えたら古いものを削除
    if (gameState.history.length > 50) {
        gameState.history.shift();
    }
}

function goBack() {
    // 履歴が2つ以上ない場合は戻れない
    if (gameState.history.length < 2) return;
    
    // 1. 現在の位置情報を履歴から削除
    gameState.history.pop();
    
    // 2. 戻るべき位置情報を取得
    const prevPosition = gameState.history.pop();
    
    // 3. ゲーム状態を戻るべき位置に強制的に更新
    gameState.currentBlock = prevPosition.block;
    gameState.currentIndex = prevPosition.index;
    
    // 4. スコア情報も復元
    if (prevPosition.totalScore !== undefined) {
        gameState.totalScore = prevPosition.totalScore;
    }
    if (prevPosition.lastChoiceScore !== undefined) {
        gameState.lastChoiceScore = prevPosition.lastChoiceScore;
    }
    
    // 5. その位置の画面を再描画
    advanceScene();
}

// ===== タイプライター効果 =====
function displayTextWithTypewriter(text) {
    clearInterval(typeInterval);
    elements.dialogueText.textContent = '';
    elements.nextIndicator.style.display = 'none';
    let index = 0;
    
    typeInterval = setInterval(() => {
        if (index < text.length) {
            elements.dialogueText.textContent += text[index];
            index++;
        } else {
            clearInterval(typeInterval);
            if (!gameState.isChoiceScene) {
                elements.nextIndicator.style.display = 'block';
            }
        }
    }, 30);
}

// ===== 画面クリック/タップ処理 =====
function handleGameScreenClick(e) {
    // UI要素のクリックは無視
    if (e.target.closest('.choice-button, .ui-button, #back-button')) return;
    if (!gameState.isPlaying || gameState.isChoiceScene) return;
    
    const currentBlockData = scenarioData[gameState.currentProtagonist][gameState.currentBlock];
    if (!currentBlockData || gameState.currentIndex > currentBlockData.length) return;
    
    const currentLineData = currentBlockData[gameState.currentIndex - 1];
    
    // タイピング中の場合は完了させる
    if (currentLineData && elements.dialogueText.textContent.length < (currentLineData.text || '').length) {
        clearInterval(typeInterval);
        elements.dialogueText.textContent = currentLineData.text;
        if (!gameState.isChoiceScene) {
            elements.nextIndicator.style.display = 'block';
        }
    } else {
        // シナリオを進行
        addToHistory();
        advanceScene();
    }
}

// ===== エンディングとセーブ =====
function showEnding() {
    const protagonistKey = gameState.currentProtagonist;
    let endingType;
    
    if (protagonistKey === '隠しシナリオ') {
        endingType = 'hidden';
    } else {
        // スコアに基づいてエンディングタイプを判定
        if (gameState.totalScore >= 6) {
            endingType = 'true';
        } else if (gameState.totalScore >= 1) {
            endingType = 'normal';
        } else {
            endingType = 'bad';
        }
    }
    
    const endingTitles = {
        true: "TRUE END",
        normal: "NORMAL END", 
        bad: "BAD END",
        hidden: "SECRET END"
    };
    
    elements.endingTitle.textContent = endingTitles[endingType];
    
    // エンディングテキストは簡潔なメッセージのみ
    elements.endingText.textContent = "プレイありがとうございました。\n\nこの物語はフィクションです！実在の人物・団体等とは関係ございません！";
    
    // TRUE ENDの場合はクリア記録を保存
    if (endingType === 'true' && protagonistKey !== '隠しシナリオ') {
        saveEndingClear(protagonistKey);
    }
    
    showScreen(elements.endingScreen);
    gameState.isPlaying = false;
}

function saveEndingClear(protagonistKey) {
    localStorage.setItem(`cleared_${protagonistKey}`, 'true');
}

function checkHiddenScenarioUnlock() {
    const shouCleared = localStorage.getItem('cleared_ショウ編') === 'true';
    const syouCleared = localStorage.getItem('cleared_しょう編') === 'true';
    
    if (shouCleared && syouCleared) {
        elements.hiddenScenarioButton.classList.remove('hidden');
    }
}

function returnToTitle() {
    gameState.isPlaying = false;
    showScreen(elements.titleScreen);
    checkHiddenScenarioUnlock();
}

function updateBackButtonVisibility() {
    const shouldShow = gameState.history.length > 1 && !gameState.isChoiceScene;
    elements.backButton.style.display = shouldShow ? 'block' : 'none';
}

// ===== 音声再生機能（フルボイス対応） =====
let currentVoiceAudio = null;

function playVoice(voiceInput) {
    // 現在の音声を停止
    stopCurrentVoice();
    
    if (!voiceInput) {
        return; // 音声IDがない場合は何もしない
    }
    
    // 配列の場合（同時再生）
    if (Array.isArray(voiceInput)) {
        currentVoiceAudio = [];
        voiceInput.forEach((voiceId, index) => {
            const audioPath = `./output_audio/${voiceId}.mp3`;
            const audio = new Audio(audioPath);
            currentVoiceAudio.push(audio);
            audio.play().catch(error => {
                console.log(`音声${index + 1}の再生に失敗: ${voiceId}`, error);
            });
        });
    } 
    // 文字列の場合（単一再生）
    else {
        const audioPath = `./output_audio/${voiceInput}.mp3`;
        currentVoiceAudio = new Audio(audioPath);
        currentVoiceAudio.play().catch(error => {
            console.log(`音声の再生に失敗: ${voiceInput}`, error);
        });
    }
}

function stopCurrentVoice() {
    if (currentVoiceAudio) {
        // 配列の場合（複数音声）
        if (Array.isArray(currentVoiceAudio)) {
            currentVoiceAudio.forEach(audio => {
                if (audio) {
                    audio.pause();
                    audio.currentTime = 0;
                }
            });
        } 
        // 単一音声の場合
        else {
            currentVoiceAudio.pause();
            currentVoiceAudio.currentTime = 0;
        }
        currentVoiceAudio = null;
    }
}