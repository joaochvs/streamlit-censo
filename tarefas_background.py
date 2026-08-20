import concurrent.futures
from dataclasses import dataclass, field
import threading
import time
import uuid

import streamlit as st


@dataclass
class Tarefa:
    inicio: float = field(default_factory=time.time)
    percentual: float = 0.0
    etapa: str = "Aguardando processamento"
    detalhe: str = "A tarefa foi adicionada à fila"
    futuro: concurrent.futures.Future | None = None


class GerenciadorTarefas:
    """Mantém tarefas vivas entre reruns e reconexões do Streamlit."""

    def __init__(self, max_workers: int = 1):
        self._executor = concurrent.futures.ThreadPoolExecutor(max_workers=max_workers)
        self._tarefas: dict[str, Tarefa] = {}
        self._lock = threading.Lock()

    def iniciar(self, funcao, arquivo_bytes: bytes) -> str:
        tarefa_id = uuid.uuid4().hex
        tarefa = Tarefa()
        with self._lock:
            self._tarefas[tarefa_id] = tarefa

        def progresso(percentual, etapa, detalhe=""):
            with self._lock:
                tarefa.percentual = float(percentual)
                tarefa.etapa = str(etapa)
                tarefa.detalhe = str(detalhe)

        tarefa.futuro = self._executor.submit(funcao, arquivo_bytes, progresso)
        return tarefa_id

    def obter(self, tarefa_id: str) -> Tarefa | None:
        with self._lock:
            return self._tarefas.get(tarefa_id)

    def remover(self, tarefa_id: str) -> None:
        """Libera o Future e seu resultado depois que a página o consumiu."""
        with self._lock:
            self._tarefas.pop(tarefa_id, None)


@st.cache_resource
def obter_gerenciador_tarefas() -> GerenciadorTarefas:
    # Limita o paralelismo para evitar que muitos usuários esgotem a memória
    # do servidor. As demais tarefas aguardam na fila sem serem perdidas.
    return GerenciadorTarefas(max_workers=1)


def acompanhar_tarefa(tarefa_id: str):
    gerenciador = obter_gerenciador_tarefas()
    tarefa = gerenciador.obter(tarefa_id)
    if tarefa is None or tarefa.futuro is None:
        raise RuntimeError("A tarefa não está mais disponível no servidor.")

    painel = st.status(tarefa.etapa, expanded=True)
    barra = st.progress(0)
    detalhes = st.empty()
    cronometro = st.empty()

    while not tarefa.futuro.done():
        decorrido = max(0, time.time() - tarefa.inicio)
        minutos, segundos = divmod(int(decorrido), 60)
        aguardando = not tarefa.futuro.running()
        etapa_exibida = "Aguardando liberação do servidor" if aguardando else tarefa.etapa
        detalhe_exibido = (
            "Outra ferramenta está sendo processada. Sua tarefa iniciará automaticamente."
            if aguardando
            else tarefa.detalhe
        )
        painel.update(label=etapa_exibida, state="running", expanded=True)
        barra.progress(min(100, max(0, int(tarefa.percentual * 100))))
        detalhes.markdown(f"**{detalhe_exibido}**" if detalhe_exibido else "")
        cronometro.markdown(f"⏱️ **Tempo decorrido** &nbsp; `{minutos:02d}:{segundos:02d}`")
        time.sleep(0.5)

    # Propaga eventuais erros do processamento para a página.
    try:
        resultado = tarefa.futuro.result()
    finally:
        # O Future retém todo o resultado em memória. Depois da entrega (ou de
        # uma falha), ele não deve permanecer no gerenciador global.
        gerenciador.remover(tarefa_id)
    decorrido = max(0, time.time() - tarefa.inicio)
    minutos, segundos = divmod(int(decorrido), 60)
    barra.progress(100)
    detalhes.markdown("✅ **Arquivo processado e pronto para download.**")
    cronometro.markdown(f"⏱️ **Tempo total** &nbsp; `{minutos:02d}:{segundos:02d}`")
    painel.update(label="Processamento concluído", state="complete", expanded=False)
    return resultado
