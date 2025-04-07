Por favor, valide o documento Excel (.xlsx) em anexo e compartilhe sua aprovação. Quaisquer ajustes necessários devem ser respondidos neste e-mail, e após os ajustes, um novo documento será gerado para validação e aprovação.

### Como validar usuários, grupos e acesso no ambiente tecnológico da organização:

1. **Aba Users**: Verifique se todos os usuários devem existir no ambiente com base em seu e-mail (coluna A) ou nome de usuário (coluna K). Usuários que foram desligados só podem estar presentes na lista se a coluna "E" estiver "Desativado", indicando que o usuário está bloqueado. Note que usuários que não pertencem a uma pessoa natural (como usuários genéricos/não nominais) não devem ser mantidos.

2. **Aba SharePointMemberships**: Verifique se todos os SharePoints devem existir no ambiente com base em seu nome (coluna A). Cada usuário com acesso ao SharePoint está listado em uma nova linha; verifique se o usuário listado (coluna D, E, F) deve ter acesso ao SharePoint indicado na mesma linha (coluna A). Além disso, verifique se a permissão de cada usuário é apropriada para seu papel (coluna G).

3. **Aba DLMemberships**: Verifique se todas as listas de distribuição devem existir no ambiente com base em seu nome (coluna A). Cada usuário que recebe mensagens de cada lista de distribuição está listado em uma nova linha; verifique se o usuário listado (coluna D, E, F) deve receber mensagens de e-mail da lista de distribuição relacionada indicada na mesma linha (coluna A). Além disso, verifique se a permissão de cada usuário é apropriada para seu papel (coluna G).

   **Descrições de Permissão**:
   - **Proprietário**: Permissão de administrador, tem controle total sobre o SharePoint/Lista de Distribuição.
   - **Membro**: Tem controle total sobre arquivos/receber mensagens, mas não pode aprovar, adicionar ou remover novos membros.
