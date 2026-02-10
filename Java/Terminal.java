import java.util.Scanner;

public class Terminal {

    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        System.out.println("--- Banco Clix ---");
        System.out.print("Número da conta: ");
        String numero = scanner.nextLine();

        System.out.print("Nome do titular: ");
        String nome = scanner.nextLine();

        System.out.print("Depósito inicial? (s/n): ");
        String resposta = scanner.nextLine();

        double deposito = 0;
        if (resposta.equalsIgnoreCase("s")) {
            System.out.print("Valor do depósito: ");
            deposito = scanner.nextDouble();
        }

        Conta conta = new Conta(numero, nome, deposito);

        System.out.println("\nConta criada!");
        System.out.println("Titular: " + conta.getNomeTitular());
        System.out.println("Conta: " + conta.getNumeroConta());
        System.out.println("Saldo: R$" + conta.getSaldo());

        int opcao = -1;
        while (opcao != 0) {
            System.out.println("\n--- Menu ---");
            System.out.println("1 - Depositar");
            System.out.println("2 - Sacar");
            System.out.println("3 - Ver Saldo");
            System.out.println("0 - Sair");
            System.out.print("Opção: ");
            opcao = scanner.nextInt();

            switch (opcao) {
                case 1:
                    System.out.print("Valor para depositar: ");
                    double valorDep = scanner.nextDouble();
                    conta.depositar(valorDep);
                    break;
                case 2:
                    System.out.print("Valor para sacar: ");
                    double valorSaq = scanner.nextDouble();
                    conta.sacar(valorSaq);
                    break;
                case 3:
                    System.out.println("Saldo atual: R$" + conta.getSaldo());
                    break;
                case 0:
                    System.out.println("Encerrando...");
                    break;
                default:
                    System.out.println("Opção inválida.");
            }
        }

        scanner.close();
    }
}
