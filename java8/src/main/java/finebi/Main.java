package finebi;

import java.nio.file.Path;
import java.nio.file.Paths;

/**
 * 命令行入口，方便不带 ElasticJob 时直接验证。
 *
 * java -jar finebi-export-runner.jar --python C:\Python313\python.exe \
 *      --work-dir D:\work\finebi-export --config D:\work\finebi-export\config.xlsx [--job 客户明细]
 */
public final class Main {

    public static void main(String[] args) {
        String python = "python";
        String workDir = ".";
        String config = null;
        String job = null;
        int maxPages = 0;

        for (int i = 0; i < args.length; i++) {
            String a = args[i];
            if ("--python".equals(a) && i + 1 < args.length) {
                python = args[++i];
            } else if ("--work-dir".equals(a) && i + 1 < args.length) {
                workDir = args[++i];
            } else if ("--config".equals(a) && i + 1 < args.length) {
                config = args[++i];
            } else if ("--job".equals(a) && i + 1 < args.length) {
                job = args[++i];
            } else if ("--max-pages".equals(a) && i + 1 < args.length) {
                maxPages = Integer.parseInt(args[++i]);
            } else if ("-h".equals(a) || "--help".equals(a)) {
                printUsage();
                return;
            } else {
                System.err.println("未知参数：" + a);
                printUsage();
                System.exit(2);
            }
        }

        Path dir = Paths.get(workDir).toAbsolutePath().normalize();
        Path cfg = config == null ? dir.resolve("config.xlsx") : Paths.get(config).toAbsolutePath();

        try {
            FineBiExportRunner runner = new FineBiExportRunner(python, dir, cfg);
            FineBiExportRunner.Result r = runner.runOnce(job, maxPages);
            // 子进程日志原样转发到 stderr，JSON 原样转发到 stdout，保持管道语义
            System.err.print(r.getStderr());
            System.out.print(r.getStdout());
            System.err.println("[java] " + r.describe());
            System.exit(r.isOk() ? 0 : 1);
        } catch (Exception e) {
            System.err.println("[java] 调用失败：" + e);
            System.exit(1);
        }
    }

    private static void printUsage() {
        System.out.println("用法：java -jar finebi-export-runner.jar [选项]");
        System.out.println("  --python    <路径>  python 可执行文件，默认 python（在 PATH 里就够）");
        System.out.println("  --work-dir  <目录>  collector.py 所在目录，默认当前目录");
        System.out.println("  --config    <文件>  配置 xlsx，默认 <work-dir>/config.xlsx");
        System.out.println("  --job       <名称>  只跑指定任务（名字或片段），不填=全部启用任务");
        System.out.println("  --max-pages <N>     每组件最多抓 N 页，调试用（0 或省略=全量）");
    }
}
