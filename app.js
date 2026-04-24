// app.js
const CONTAINER_L = 12000;
const CONTAINER_W = 2300;
const CONTAINER_H = 2400;

const BOX_TYPES = [
    { name: 'Case Type A', l: 1100, w: 1100, h: 1000, color: 0x3b82f6 },
    { name: 'Case Type B', l: 1200, w: 1100, h: 800,  color: 0x10b981 },
    { name: 'Case Type C', l: 800,  w: 1100, h: 1200, color: 0xf59e0b },
    { name: 'Case Type D', l: 1400, w: 1100, h: 1100, color: 0x6366f1 },
    { name: 'Case Type E', l: 1000, w: 1100, h: 1100, color: 0xec4899 }
];

let pendingBoxes = [];
let packedList = [];
let currentMonth = 0;
let scene, camera, renderer, controls;
let containerBoxesGroup = new THREE.Group();

function initThreeJS() {
    const container = document.getElementById('canvas-container');
    scene = new THREE.Scene();
    scene.background = null; // Let CSS handle gradient background

    camera = new THREE.PerspectiveCamera(45, container.clientWidth / container.clientHeight, 1, 100000);
    // Initial camera position for generic 3D view
    camera.position.set(15000, 8000, 15000);

    renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true });
    renderer.setSize(container.clientWidth, container.clientHeight);
    container.appendChild(renderer.domElement);

    controls = new THREE.OrbitControls(camera, renderer.domElement);
    controls.enableDamping = true;
    controls.dampingFactor = 0.05;
    controls.target.set(CONTAINER_W/2, CONTAINER_H/2, CONTAINER_L/2);

    // Grid helper mapped to floor
    const gridHelper = new THREE.GridHelper(20000, 20, 0x334155, 0x1e293b);
    gridHelper.position.set(CONTAINER_W/2, 0, CONTAINER_L/2);
    scene.add(gridHelper);

    // Lights
    const ambientLight = new THREE.AmbientLight(0xffffff, 0.7);
    scene.add(ambientLight);
    
    const dirLight = new THREE.DirectionalLight(0xffffff, 0.8);
    dirLight.position.set(10000, 20000, 10000);
    scene.add(dirLight);

    // Container wireframe bounds
    const geo = new THREE.BoxGeometry(CONTAINER_W, CONTAINER_H, CONTAINER_L);
    const edges = new THREE.EdgesGeometry(geo);
    const mat = new THREE.LineBasicMaterial({ color: 0x60a5fa, transparent: true, opacity: 0.5 });
    const wireframe = new THREE.LineSegments(edges, mat);
    wireframe.position.set(CONTAINER_W/2, CONTAINER_H/2, CONTAINER_L/2);
    scene.add(wireframe);

    // A subtle floor for the container
    const floorGeo = new THREE.PlaneGeometry(CONTAINER_W, CONTAINER_L);
    const floorMat = new THREE.MeshBasicMaterial({ color: 0x3b82f6, transparent: true, opacity: 0.1, side: THREE.DoubleSide });
    const floorMesh = new THREE.Mesh(floorGeo, floorMat);
    floorMesh.rotation.x = Math.PI / 2;
    floorMesh.position.set(CONTAINER_W/2, 0, CONTAINER_L/2);
    scene.add(floorMesh);

    scene.add(containerBoxesGroup);

    window.addEventListener('resize', () => {
        camera.aspect = container.clientWidth / container.clientHeight;
        camera.updateProjectionMatrix();
        renderer.setSize(container.clientWidth, container.clientHeight);
    });

    animate();
}

function animate() {
    requestAnimationFrame(animate);
    controls.update();
    renderer.render(scene, camera);
}

// GUI and Logic Interaction
document.getElementById('btn-simulate').addEventListener('click', () => {
    currentMonth = 1;
    packedList = [];
    pendingBoxes = generateMonthBoxes();
    runSimulation();
    
    document.getElementById('btn-next-month').disabled = false;
    document.getElementById('btn-download').disabled = false;
});

document.getElementById('btn-next-month').addEventListener('click', () => {
    currentMonth++;
    pendingBoxes = pendingBoxes.concat(generateMonthBoxes());
    runSimulation();
});

function generateMonthBoxes() {
    let boxes = [];
    // Number of boxes generating each month: usually ~10 to 30.
    const count = Math.floor(Math.random() * 21) + 10; 
    
    const minWt = parseInt(document.getElementById('weight-min').value);
    const maxWt = parseInt(document.getElementById('weight-max').value);
    
    for(let i=0; i<count; i++) {
        let type = BOX_TYPES[Math.floor(Math.random() * BOX_TYPES.length)];
        // Skewing towards lower weight randomly to mimic realistic variations avoiding ultra-heavy blockings continually
        let randWt = minWt + (Math.random() * Math.random() * Math.random()) * (maxWt - minWt);
        boxes.push({
            id: `M${currentMonth}-${(i+1).toString().padStart(3, '0')}`,
            month: currentMonth,
            name: type.name,
            l: type.l, w: type.w, h: type.h,
            color: type.color,
            weight: Math.round(randWt)
        });
    }
    return boxes;
}

class RowPacker2D {
    constructor(rowWidth, containerL, containerH) {
        this.width = rowWidth;
        this.L = containerL;
        this.H = containerH;
        this.freeSpaces = [{x: 0, z: 0, l: containerL, h: containerH}];
    }

    pack(box) {
        if (box.w > this.width) return null;

        this.freeSpaces.sort((a, b) => {
            if (a.z !== b.z) return a.z - b.z; // Lowest height first
            return a.x - b.x; // Back of container first
        });

        for (let i = 0; i < this.freeSpaces.length; i++) {
            let space = this.freeSpaces[i];
            
            if (box.l <= space.l && box.h <= space.h) {
                this.freeSpaces.splice(i, 1);
                
                let wRem = space.l - box.l;
                let hRem = space.h - box.h;
                
                if (wRem > 0 || hRem > 0) {
                    let vSplit1 = {x: space.x + box.l, z: space.z, l: wRem, h: space.h};
                    let vSplit2 = {x: space.x, z: space.z + box.h, l: box.l, h: hRem};
                    
                    let hSplit1 = {x: space.x + box.l, z: space.z, l: wRem, h: box.h};
                    let hSplit2 = {x: space.x, z: space.z + box.h, l: space.l, h: hRem};
                    
                    if (wRem > hRem) {
                        if (wRem > 0 && box.h > 0) this.freeSpaces.push(hSplit1);
                        if (space.l > 0 && hRem > 0) this.freeSpaces.push(hSplit2);
                    } else {
                        if (wRem > 0 && space.h > 0) this.freeSpaces.push(vSplit1);
                        if (box.l > 0 && hRem > 0) this.freeSpaces.push(vSplit2);
                    }
                }
                return {x: space.x, z: space.z};
            }
        }
        return null;
    }
}

class DualRowPacker {
    constructor(c_L, c_W, c_H, maxW) {
        this.L = c_L;
        this.W = c_W;
        this.H = c_H;
        this.maxWeightTarget = maxW;
        
        this.leftRow = new RowPacker2D(1150, c_L, c_H);
        this.rightRow = new RowPacker2D(1150, c_L, c_H);
        
        this.packedBoxes = [];
        this.currentWeight = 0;
        this.lastRowIdx = 0;
    }

    pack(box) {
        if (this.currentWeight + box.weight > this.maxWeightTarget) {
            return false;
        }

        // Alternate rows to balance them visually
        let firstIdx = 1 - this.lastRowIdx;
        let secondIdx = this.lastRowIdx;
        
        let rowObjects = [
            { packer: this.leftRow, y: 0 },
            { packer: this.rightRow, y: 1150 }
        ];

        let rows = [ rowObjects[firstIdx], rowObjects[secondIdx] ];
        
        for (let idx=0; idx<rows.length; idx++) {
            let row = rows[idx];
            let pos = row.packer.pack(box);
            if (pos) {
                box.x = pos.x;
                box.y = row.y + (1150 - box.w) / 2;
                box.z = pos.z;
                
                this.packedBoxes.push(box);
                this.currentWeight += box.weight;
                
                // Switch preference if we placed it
                if (idx === 0) {
                    this.lastRowIdx = firstIdx;
                } else {
                    this.lastRowIdx = secondIdx;
                }
                
                return true;
            }
        }
        return false;
    }
}

function runSimulation() {
    const maxWeight = parseFloat(document.getElementById('max-weight').value);
    const targetMax = parseFloat(document.getElementById('target-max').value) / 100;
    let upperLimit = maxWeight * targetMax;
    
    const packer = new DualRowPacker(CONTAINER_L, CONTAINER_W, CONTAINER_H, upperLimit);
    
    let allBoxes = [...packedList, ...pendingBoxes];
    
    // Optimally try to place Heaviest and Largest Volume first for tighter packing
    allBoxes.sort((a,b) => {
        if (a.month !== b.month) return a.month - b.month;
        let areaA = a.l * a.h;
        let areaB = b.l * b.h;
        if (areaB !== areaA) return areaB - areaA;
        return a.weight - b.weight; // Prioritize lighter boxes if same volume to fit more
    });

    let toKeepPending = [];
    
    for(let box of allBoxes) {
        if(!packer.pack(box)) {
            toKeepPending.push(box);
        }
    }

    packedList = packer.packedBoxes;
    pendingBoxes = toKeepPending;

    updateUI(maxWeight);
    drawBoxes(packedList);
}

function updateUI(maxWeight) {
    let weight = packedList.reduce((sum, b) => sum + b.weight, 0);
    document.getElementById('stat-month').innerText = `${currentMonth} ヵ月目`;
    document.getElementById('stat-boxes').innerText = packedList.length;
    document.getElementById('stat-weight').innerText = weight.toLocaleString();
    
    let pct = ((weight / maxWeight) * 100);
    let pctEl = document.getElementById('stat-weight-pct');
    pctEl.innerText = pct.toFixed(1) + '%';
    
    // Coloring logic for validation target percentages
    if(pct >= 70 && pct <= 85) {
        pctEl.style.color = 'var(--success)';
    } else if (pct > 85) {
        pctEl.style.color = '#ef4444';
    } else {
        pctEl.style.color = 'var(--warning)';
    }
    
    document.getElementById('stat-pending').innerText = pendingBoxes.length;
    
    // Update Sidebar List Table
    const tbody = document.querySelector('#loaded-boxes-table tbody');
    tbody.innerHTML = '';
    
    // Sort logic for visual display from back to front
    let displayList = [...packedList].sort((a, b) => a.x - b.x);
    
    document.getElementById('list-count').innerText = displayList.length;
    
    displayList.forEach(box => {
        let tr = document.createElement('tr');
        tr.dataset.id = box.id;
        tr.innerHTML = `
            <td>${box.month}</td>
            <td>${box.id}</td>
            <td>${box.l}x${box.w}x${box.h}</td>
            <td>${box.weight.toLocaleString()}</td>
        `;
        
        tr.addEventListener('mouseenter', () => highlightBox(box.id, true));
        tr.addEventListener('mouseleave', () => highlightBox(box.id, false));
        tbody.appendChild(tr);
    });
}

function highlightBox(id, isHighlight) {
    containerBoxesGroup.children.forEach(mesh => {
        if (mesh.userData && mesh.userData.id === id) {
            if (isHighlight) {
                mesh.material.emissive.setHex(0x3b82f6);
                mesh.material.emissiveIntensity = 0.5;
            } else {
                mesh.material.emissive.setHex(0x000000);
            }
        }
    });
}

function drawBoxes(boxes) {
    // Clear old geometry references thoroughly
    while(containerBoxesGroup.children.length > 0) {
        let child = containerBoxesGroup.children[0];
        child.geometry.dispose();
        child.material.dispose();
        // remove edges
        if(child.children.length > 0) {
            child.children[0].geometry.dispose();
            child.children[0].material.dispose();
        }
        containerBoxesGroup.remove(child);
    }

    boxes.forEach(box => {
        const geo = new THREE.BoxGeometry(box.w, box.h, box.l);
        const mat = new THREE.MeshPhongMaterial({ 
            color: box.color,
            transparent: true,
            opacity: 0.9,
            shininess: 30
        });
        const mesh = new THREE.Mesh(geo, mat);
        
        // Wireframe mapped to edges for clarity
        const edgeGeo = new THREE.EdgesGeometry(geo);
        const edgeMat = new THREE.LineBasicMaterial({ color: 0xffffff, linewidth: 2, transparent: true, opacity: 0.4 });
        const edges = new THREE.LineSegments(edgeGeo, edgeMat);
        mesh.add(edges);

        // Three.js mappings:
        // box.x specifies the Depth penetration into the container (z-axis in standard 3D usually, but we set container length to Z here)
        // Set coordinates using our mapping: container Width (X), container Height (Y), container Length (Z)
        mesh.position.set(
            box.y + box.w/2, 
            box.z + box.h/2, 
            box.x + box.l/2 
        );
        
        mesh.userData = box;
        containerBoxesGroup.add(mesh);
    });
}

// Raycaster for Hovering Data Tooltip
const raycaster = new THREE.Raycaster();
const mouse = new THREE.Vector2();
const containerEl = document.getElementById('canvas-container');

containerEl.addEventListener('mousemove', (e) => {
    const rect = renderer.domElement.getBoundingClientRect();
    mouse.x = ((e.clientX - rect.left) / rect.width) * 2 - 1;
    mouse.y = -((e.clientY - rect.top) / rect.height) * 2 + 1;
    
    raycaster.setFromCamera(mouse, camera);
    const intersects = raycaster.intersectObjects(containerBoxesGroup.children, false);
    
    const tooltip = document.getElementById('tooltip');
    
    if (intersects.length > 0) {
        const box = intersects[0].object.userData;
        if(box.id) {
            tooltip.innerHTML = `
                <h4>${box.name} (${box.id})</h4>
                <p><strong>重量:</strong> ${box.weight.toLocaleString()} kg</p>
                <p><strong>サイズ:</strong> ${box.l}x${box.w}x${box.h} mm</p>
                <p><strong>搬入月:</strong> ${box.month}ヵ月目</p>
            `;
            tooltip.style.left = e.clientX + 15 + 'px';
            tooltip.style.top = e.clientY + 15 + 'px';
            tooltip.classList.remove('hidden');
            document.body.style.cursor = 'pointer';
            return;
        }
    }
    tooltip.classList.add('hidden');
    document.body.style.cursor = 'default';
});

containerEl.addEventListener('mouseleave', () => {
    document.getElementById('tooltip').classList.add('hidden');
});

// View Controls Logic
document.querySelectorAll('.view-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
        document.querySelectorAll('.view-btn').forEach(b => b.classList.remove('active'));
        e.target.classList.add('active');
        
        const view = e.target.dataset.view;
        if(view === '3d') {
            camera.position.set(15000, 8000, 15000);
        } else if (view === 'top') {
            camera.position.set(CONTAINER_W/2, 20000, CONTAINER_L/2);
        } else if (view === 'side') {
            camera.position.set(20000, CONTAINER_H/2, CONTAINER_L/2);
        }
        controls.target.set(CONTAINER_W/2, CONTAINER_H/2, CONTAINER_L/2);
    });
});

// CSV Export Logic
document.getElementById('btn-download').addEventListener('click', () => {
    let csv = 'ID,月度,ケース名,長さ(mm),幅(mm),高さ(mm),重量(kg)\n';
    
    // Sort logically for export file: by month then by placement Z (lengthwise)
    let exportList = [...packedList].sort((a,b) => {
        if(a.month !== b.month) return a.month - b.month;
        return a.x - b.x;
    });

    exportList.forEach(b => {
        csv += `${b.id},${b.month},${b.name},${b.l},${b.w},${b.h},${b.weight}\n`;
    });
    
    let filename = `vanning_result_month${currentMonth}.csv`;

    if (window.pywebview && window.pywebview.api) {
        // Native Desktop App mode
        window.pywebview.api.save_csv(filename, csv).then((saved) => {
            if (saved) {
                console.log("CSV saved successfully via python API.");
            }
        }).catch(err => {
            console.error("Failed to save via python API", err);
            alert('保存に失敗しました。');
        });
    } else {
        // Fallback for running locally in Browser
        const blob = new Blob([new Uint8Array([0xEF, 0xBB, 0xBF]), csv], {type: 'text/csv;charset=utf-8;'});
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        a.click();
        URL.revokeObjectURL(url);
    }
});

// Initializing the context
setTimeout(initThreeJS, 100); // small delay to ensure DOM is fully ready
