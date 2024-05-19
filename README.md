# **System Clear**

#### *Ferramenta para limpeza básica e desativação de Inicialização Rápida, visando melhorar o desempenho do sistema.*

##### Detalhadamente, o System Clear realiza todas essas tarefas:
- Remove arquivos temporários acumulados na pasta Temp do usuário.
- Remove arquivos temporários acumulados na pasta Temp do sistema.
- Remove permanentemente os arquivos descartados pelo usuário.
- Remove o cache e outros arquivos temporários dos navegadores.
- Remove arquivos de atualizações do Windows que não são mais necessários.
- Remove registros de atualizações anteriores do Windows.
- Remove arquivos desnecessários e temporários através da ferramenta Cleanmgr.
- Desabilita o recurso de Inicialização Rápida para melhorar o desempenho do sistema.
- Verifica se há documentos não salvos em programas como Notepad, Word, Excel e PowerPoint, alertando o usuário antes de realizar a reinicialização.
- Reinicialização opcional do sistema: Pergunta ao usuário se deseja reiniciar o sistema após completar as tarefas de limpeza.

#### Instalação
1. Abra o PowerShell como Administrador:
    - Pressione Win + X (ou clique com o botão direito do mouse no menu Iniciar) e selecione Windows PowerShell (Admin).
2. Ajuste a Política de Execução:
```
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
```
3. Execute o Script (start.bat) como administrador.

#### License
Ana Carla Callegarim Della Giacomo [anaccdg]
<div style="display: flex; align-items: center;">  <a href="https://github.com/acdgiacomo" style="display: flex;">  <img src="https://cdn-icons-png.flaticon.com/256/25/25231.png" alt="GitHub" width="50" height="50"> </a> </div>
