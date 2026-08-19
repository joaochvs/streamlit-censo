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

    def __init__(self, max_workers: int = 2):
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


@st.cache_resource
def obter_gerenciador_tarefas() -> GerenciadorTarefas:
    # Limita o paralelismo para evitar que muitos usuários esgotem a memória
    # do servidor. As demais tarefas aguardam na fila sem serem perdidas.
    return GerenciadorTarefas(max_workers=2)


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
        painel.update(label=tarefa.etapa, state="running", expanded=True)
        barra.progress(min(100, max(0, int(tarefa.percentual * 100))))
        detalhes.markdown(f"**{tarefa.detalhe}**" if tarefa.detalhe else "")
        cronometro.markdown(f"⏱️ **Tempo decorrido** &nbsp; `{minutos:02d}:{segundos:02d}`")
        time.sleep(0.5)

    # Propaga eventuais erros do processamento para a página.
    resultado = tarefa.futuro.result()
    decorrido = max(0, time.time() - tarefa.inicio)
    minutos, segundos = divmod(int(decorrido), 60)
    barra.progress(100)
    detalhes.markdown("✅ **Arquivo processado e pronto para download.**")
    cronometro.markdown(f"⏱️ **Tempo total** &nbsp; `{minutos:02d}:{segundos:02d}`")
    painel.update(label="Processamento concluído", state="complete", expanded=False)
    return resultado
