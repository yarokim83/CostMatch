document.addEventListener('DOMContentLoaded', () => {
    const dropZones = [
        { zone: document.getElementById('drop-zone-1'), input: document.getElementById('file1'), info: document.getElementById('file-info-1') },
        { zone: document.getElementById('drop-zone-2'), input: document.getElementById('file2'), info: document.getElementById('file-info-2') }
    ];

    dropZones.forEach(({ zone, input, info }) => {
        zone.addEventListener('click', () => input.click());

        input.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                info.textContent = e.target.files[0].name;
                zone.classList.add('active');
            }
        });

        zone.addEventListener('dragover', (e) => {
            e.preventDefault();
            zone.classList.add('dragover');
        });

        zone.addEventListener('dragleave', () => {
            zone.classList.remove('dragover');
        });

        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            zone.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) {
                input.files = e.dataTransfer.files;
                info.textContent = e.dataTransfer.files[0].name;
            }
        });
    });

    document.getElementById('compare-btn').addEventListener('click', async () => {
        const file1 = document.getElementById('file1').files[0];
        const file2 = document.getElementById('file2').files[0];

        if (!file1 || !file2) {
            alert('두 파일을 모두 선택해주세요.');
            return;
        }

        const formData = new FormData();
        formData.append('file1', file1);
        formData.append('file2', file2);

        const btn = document.getElementById('compare-btn');
        btn.textContent = '분석 중...';
        btn.disabled = true;

        try {
            const response = await fetch('/compare', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            document.getElementById('result-section').classList.remove('hidden');
            const output = document.getElementById('output');

            if (response.ok) {
                output.innerHTML = `
                    <p>✅ 파일 로드 성공</p>
                    <p><strong>File 1 컬럼:</strong> ${result.file1_columns.join(', ')}</p>
                    <p><strong>File 2 컬럼:</strong> ${result.file2_columns.join(', ')}</p>
                    <p><em>실제 비교 로직은 컬럼 확인 후 구현됩니다.</em></p>
                `;
            } else {
                output.innerHTML = `<p style="color: red;">❌ 오류: ${result.error}</p>`;
            }
        } catch (error) {
            console.error(error);
            alert('오류가 발생했습니다.');
        } finally {
            btn.textContent = '비교 분석 시작';
            btn.disabled = false;
        }
    });
});
