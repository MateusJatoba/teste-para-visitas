function adicionarVisitante() {
            const container = document.getElementById('visitantes');
            const index = container.children.length;

            const novoGrupo = document.createElement('div');
            novoGrupo.className = 'form-group';
            
            novoGrupo.innerHTML = `
                <label for="nome_visitante_${index}">Nome Completo do Visitante:</label>
                <input type="text" name="nome_visitante[]" id="nome_visitante_${index}" required>

                <label for="cargo_visitante_${index}">Cargo do Visitante:</label>
                <input type="text" name="cargo_visitante[]" id="cargo_visitante_${index}" required>
            `;

            container.appendChild(novoGrupo);
        }