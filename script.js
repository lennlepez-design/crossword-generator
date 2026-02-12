let wordsFromExcel = [];
const SIZE = 45; 
let timerInterval;
let seconds = 0;
let currentDirection = 'across';

document.getElementById('excelUpload').addEventListener('change', function(e) {
    const reader = new FileReader();
    reader.onload = function(event) {
        const data = new Uint8Array(event.target.result);
        const wb = XLSX.read(data, {type: 'array'});
        const rawData = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header: 1});
        wordsFromExcel = rawData.filter(r => r[0] && r[1]).map(r => ({
            word: String(r[0]).trim().toUpperCase().replace(/^(DE\s+|HET\s+|EEN\s+)/g, ""),
            clue: String(r[1]).trim()
        }));
        document.getElementById('status').innerText = `✅ ${wordsFromExcel.length} woorden geladen`;
    };
    reader.readAsArrayBuffer(e.target.files[0]);
});

function generatePuzzle() {
    if (!wordsFromExcel.length) return alert("Upload eerst Excel!");
    clearInterval(timerInterval); seconds = 0;
    timerInterval = setInterval(() => {
        seconds++;
        let m = Math.floor(seconds / 60).toString().padStart(2, '0');
        let s = (seconds % 60).toString().padStart(2, '0');
        document.getElementById('timer').innerText = `${m}:${s}`;
    }, 1000);

    let best = attemptBuild();
    renderGrid(best.grid, best.placedWords);
    renderSolutionGrid(best.grid); // Generates the correction key
}

function attemptBuild() {
    let grid = Array(SIZE).fill().map(() => Array(SIZE).fill(null));
    let placedWords = [];
    let letterMap = {};
    let pool = [...wordsFromExcel].sort((a,b) => b.word.length - a.word.length);
    const first = pool.shift();
    if(first) placeWord(grid, first, 15, 10, 'across', placedWords, letterMap);
    
    let retryPool = [];
    for (let item of pool) {
        let res = findBestSpot(grid, item.word, letterMap);
        if (res) placeWord(grid, item, res.r, res.c, res.dir, placedWords, letterMap);
        else retryPool.push(item);
    }
    for (let item of retryPool) {
        let placed = false;
        for (let r = 2; r < SIZE - 2 && !placed; r += 2) {
            for (let c = 2; c < SIZE - item.word.length - 2 && !placed; c++) {
                if (canPlace(grid, item.word, r, c, 'across')) {
                    placeWord(grid, item, r, c, 'across', placedWords, letterMap);
                    placed = true;
                }
            }
        }
    }
    return { grid, placedWords };
}

function findBestSpot(grid, word, letterMap) {
    let possibilities = [];
    for(let i=0; i<word.length; i++) {
        let letter = word[i];
        if(letterMap[letter]) {
            for(let pos of letterMap[letter]) {
                let dirs = [{r: pos.r, c: pos.c-i, dir:'across'}, {r: pos.r-i, c: pos.c, dir:'down'}];
                for(let d of dirs) {
                    if(canPlace(grid, word, d.r, d.c, d.dir)) {
                        possibilities.push({ ...d, score: countIntersections(grid, word, d.r, d.c, d.dir) });
                    }
                }
            }
        }
    }
    return possibilities.length > 0 ? possibilities.sort((a,b) => b.score - a.score)[0] : null;
}

function canPlace(grid, word, r, c, dir) {
    if (r < 1 || c < 1 || (dir === 'across' && c + word.length >= SIZE-1) || (dir === 'down' && r + word.length >= SIZE-1)) return false;
    for (let i = 0; i < word.length; i++) {
        let curR = r + (dir === 'down' ? i : 0);
        let curC = c + (dir === 'across' ? i : 0);
        if (grid[curR][curC] && grid[curR][curC].letter !== word[i]) return false;
        if (!grid[curR][curC]) {
            if (dir === 'across') {
                if (grid[curR-1][curC] || grid[curR+1][curC]) return false;
                if (i === 0 && grid[curR][curC-1]) return false;
                if (i === word.length-1 && grid[curR][curC+1]) return false;
            } else {
                if (grid[curR][curC-1] || grid[curR][curC+1]) return false;
                if (i === 0 && grid[curR-1][curC]) return false;
                if (i === word.length-1 && grid[curR+1][curC]) return false;
            }
        }
    }
    return true;
}

function countIntersections(grid, word, r, c, dir) {
    let score = 0;
    for(let i=0; i<word.length; i++) {
        let curR = r + (dir === 'down' ? i : 0);
        let curC = c + (dir === 'across' ? i : 0);
        if(grid[curR][curC]) score++;
    }
    return score;
}

function placeWord(grid, info, r, c, dir, list, letterMap) {
    const num = list.length + 1;
    let cells = [];
    for (let i = 0; i < info.word.length; i++) {
        let curR = r + (dir === 'down' ? i : 0);
        let curC = c + (dir === 'across' ? i : 0);
        cells.push(`${curR}-${curC}`);
        if (!grid[curR][curC]) {
            grid[curR][curC] = { letter: info.word[i], num: i === 0 ? num : null, words: [] };
            if (!letterMap[info.word[i]]) letterMap[info.word[i]] = [];
            letterMap[info.word[i]].push({ r: curR, c: curC });
        } else if (i === 0) grid[curR][curC].num = num;
        grid[curR][curC].words.push({ num, dir, word: info.word, clue: info.clue, cells: cells });
    }
    list.push({ ...info, num, dir, cells });
}

function renderGrid(grid, placedWords) {
    const container = document.getElementById('grid-container');
    container.innerHTML = '';
    container.style.gridTemplateColumns = `repeat(${SIZE}, 22px)`;

    for (let r = 0; r < SIZE; r++) {
        for (let c = 0; c < SIZE; c++) {
            const cell = document.createElement('div');
            cell.className = 'cell' + (grid[r][c] ? '' : ' black');
            cell.id = `cell-${r}-${c}`;
            if (grid[r][c]) {
                if (grid[r][c].num) cell.innerHTML += `<span class="num">${grid[r][c].num}</span>`;
                const input = document.createElement('input');
                input.maxLength = 1;
                input.dataset.ans = grid[r][c].letter;
                input.dataset.r = r; input.dataset.c = c;
                
                input.addEventListener('focus', () => {
                    let activeWord = grid[r][c].words.find(w => w.dir === currentDirection) || grid[r][c].words[0];
                    currentDirection = activeWord.dir;
                    highlightWord(activeWord, placedWords);
                    document.querySelectorAll('.cell').forEach(ce => ce.classList.remove('active-focus'));
                    cell.classList.add('active-focus');
                    input.select(); // Highlight existing letter on focus
                });
                
                input.addEventListener('input', (e) => {
                    e.target.value = e.target.value.toUpperCase();
                    validateWord(grid[r][c].words);
                    if (e.target.value) handleNavigation(e.target);
                });

                input.addEventListener('keydown', (e) => {
                    if (e.key === "Backspace" && !e.target.value) handleBack(e.target);
                });

                cell.appendChild(input);
            }
            container.appendChild(cell);
        }
    }
    renderClues(placedWords);
    document.getElementById('stats').innerText = `Totaal: ${placedWords.length}`;
}

function validateWord(words) {
    words.forEach(w => {
        let currentString = "";
        let inputElements = [];
        w.cells.forEach(id => {
            const el = document.querySelector(`#cell-${id} input`);
            currentString += el.value;
            inputElements.push(el);
        });

        if (currentString.length === w.word.length) {
            const isCorrect = currentString === w.word;
            inputElements.forEach(el => {
                el.classList.add(isCorrect ? 'correct' : 'wrong');
                el.classList.remove(isCorrect ? 'wrong' : 'correct');
            });
        } else {
            inputElements.forEach(el => el.classList.remove('correct', 'wrong'));
        }
    });
}

function handleNavigation(input) {
    const r = parseInt(input.dataset.r), c = parseInt(input.dataset.c);
    const nextR = r + (currentDirection === 'down' ? 1 : 0);
    const nextC = c + (currentDirection === 'across' ? 1 : 0);
    const nextInput = document.querySelector(`#cell-${nextR}-${nextC} input`);
    
    if (nextInput) {
        nextInput.focus();
        nextInput.select(); // Enable retyping by selecting existing content
    }
}

function handleBack(input) {
    const r = parseInt(input.dataset.r), c = parseInt(input.dataset.c);
    const prevR = r - (currentDirection === 'down' ? 1 : 0);
    const prevC = c - (currentDirection === 'across' ? 1 : 0);
    const prevInput = document.querySelector(`#cell-${prevR}-${prevC} input`);
    if (prevInput) {
        prevInput.focus();
        prevInput.select();
    }
}

function highlightWord(wordInfo, placedWords) {
    document.querySelectorAll('.cell').forEach(c => c.classList.remove('highlighted'));
    const fullInfo = placedWords.find(w => w.num === wordInfo.num && w.dir === wordInfo.dir);
    fullInfo.cells.forEach(id => document.getElementById(`cell-${id}`).classList.add('highlighted'));
    document.getElementById('active-clue-banner').innerText = `${fullInfo.num} ${fullInfo.dir === 'across' ? 'Horizontaal' : 'Verticaal'}: ${fullInfo.clue}`;
    document.querySelectorAll('li').forEach(l => l.classList.remove('active-hint'));
    const hintEl = document.getElementById(`hint-${fullInfo.dir}-${fullInfo.num}`);
    if (hintEl) hintEl.classList.add('active-hint');
}

function renderClues(placed) {
    const hor = document.getElementById('across-clues'), ver = document.getElementById('down-clues');
    hor.innerHTML = ''; ver.innerHTML = '';
    placed.sort((a,b) => a.num - b.num).forEach(w => {
        const li = document.createElement('li');
        li.id = `hint-${w.dir}-${w.num}`;
        li.innerHTML = `<b>${w.num}.</b> ${w.clue}`;
        li.onclick = () => {
            currentDirection = w.dir;
            const target = document.querySelector(`#cell-${w.cells[0]} input`);
            target.focus();
            target.select();
        };
        (w.dir === 'across' ? hor : ver).appendChild(li);
    });
}

function renderSolutionGrid(grid) {
    let solutionWrapper = document.getElementById('solution-wrapper');
    if (!solutionWrapper) {
        solutionWrapper = document.createElement('div');
        solutionWrapper.id = 'solution-wrapper';
        solutionWrapper.innerHTML = `
            <div style="margin-top: 50px; border-top: 2px dashed #ccc; padding-top: 20px;">
                <h2 style="color: var(--dark); text-align: center;">Correctiesleutel (Oplossing)</h2>
                <div id="solution-container"></div>
            </div>`;
        document.querySelector('.paper').appendChild(solutionWrapper);
    }

    const container = document.getElementById('solution-container');
    container.innerHTML = '';
    container.style.display = 'grid';
    container.style.gridTemplateColumns = `repeat(${SIZE}, 12px)`;
    container.style.gap = '1px';
    container.style.background = '#ddd';
    container.style.width = 'fit-content';
    container.style.margin = '20px auto';
    container.className = 'no-print'; // Optional: remove if you want solution on print

    for (let r = 0; r < SIZE; r++) {
        for (let c = 0; c < SIZE; c++) {
            const cell = document.createElement('div');
            cell.style.width = '12px';
            cell.style.height = '12px';
            cell.style.fontSize = '8px';
            cell.style.display = 'flex';
            cell.style.alignItems = 'center';
            cell.style.justifyContent = 'center';
            cell.style.fontWeight = 'bold';
            
            if (grid[r][c]) {
                cell.style.background = 'white';
                cell.innerText = grid[r][c].letter;
                cell.style.color = 'var(--primary)';
            } else {
                cell.style.background = '#eee';
            }
            container.appendChild(cell);
        }
    }
}

function checkPassword() {
    if (prompt("Wachtwoord?") === "12345") {
        document.querySelectorAll('.paper input').forEach(i => { 
            i.value = i.dataset.ans; 
            i.classList.add('correct');
        });
    }
}