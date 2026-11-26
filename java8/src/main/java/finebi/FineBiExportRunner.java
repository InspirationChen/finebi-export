package finebi;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Java 8 侧的调用壳：把 Python 采集器当成一个「可执行的 Job」来调。
 *
 * 设计要点（都是为了让 ElasticJob / Quartz / Jenkins 好接）：
 *   1. 一次调用只跑一轮，跑完进程退出，不依赖任何常驻服务；
 *   2. 结果从 stdout 取 JSON，日志从 stderr 取，两者分离；
 *   3. 退出码语义：0 = 全部成功，1 = 有任务失败，2 = 配置/参数错误；
 *   4. stdout / stderr 并发消费，避免管道写满导致子进程死锁（Windows 上尤其明显）；
 *   5. 显式指定 UTF-8，避免 Windows 控制台 GBK 把中文日志搞成乱码。
 *
 * 用法（Java）：
 *   FineBiExportRunner r = new FineBiExportRunner(
 *           "C:\\\\Python313\\\\python.exe",
 *           Paths.get("D:\\\\work\\\\finebi-export"),
 *           Paths.get("D:\\\\work\\\\finebi-export\\\\config.xlsx"));
 *   FineBiExportRunner.Result result = r.runOnce(null);   // null = 跑配置里所有启用的任务
 *   if (!result.isOk()) { throw new RuntimeException(result.getStderr()); }
 */
public final class FineBiExportRunner {

    private final String pythonExe;
    private final Path workDir;
    private final Path configFile;
    private final long timeoutSeconds;

    public FineBiExportRunner(String pythonExe, Path workDir, Path configFile) {
        this(pythonExe, workDir, configFile, 1800L);   // 默认 30 分钟上限
    }

    public FineBiExportRunner(String pythonExe, Path workDir, Path configFile, long timeoutSeconds) {
        this.pythonExe = pythonExe;
        this.workDir = workDir;
        this.configFile = configFile;
        this.timeoutSeconds = timeoutSeconds;
    }

    /** 跑配置里全部启用任务。 */
    public Result runOnce() throws IOException, InterruptedException {
        return runOnce(null, 0);
    }

    /**
     * 跑一轮。
     *
     * @param jobName 只跑指定任务（名字或名字片段）；null = 全部启用任务
     */
    public Result runOnce(String jobName) throws IOException, InterruptedException {
        return runOnce(jobName, 0);
    }

    /**
     * 跑一轮。
     *
     * @param jobName  只跑指定任务（名字或片段）；null = 全部启用任务
     * @param maxPages 每组件最多抓几页；0 = 全量。调试时给 1~2 能几秒出结果
     */
    public Result runOnce(String jobName, int maxPages) throws IOException, InterruptedException {
        List<String> cmd = new ArrayList<String>();
        cmd.add(pythonExe);
        cmd.add(workDir.resolve("collector.py").toString());
        cmd.add("--config");
        cmd.add(configFile.toAbsolutePath().toString());
        cmd.add("--once");
        cmd.add("--json");
        if (jobName != null && !jobName.trim().isEmpty()) {
            cmd.add("--job");
            cmd.add(jobName.trim());
        }
        if (maxPages > 0) {
            cmd.add("--max-pages");
            cmd.add(String.valueOf(maxPages));
        }

        ProcessBuilder pb = new ProcessBuilder(cmd);
        pb.directory(workDir.toFile());
        pb.redirectErrorStream(false);
        // Windows 下这两行很关键：不加的话 Python 输出按 GBK 编码，Java 按 UTF-8 读会乱码
        pb.environment().put("PYTHONUTF8", "1");
        pb.environment().put("PYTHONIOENCODING", "utf-8");

        long startedAt = System.currentTimeMillis();
        Process process = pb.start();

        StringBuilder stdout = new StringBuilder();
        StringBuilder stderr = new StringBuilder();
        Thread outPump = pump(process.getInputStream(), stdout, "finebi-stdout");
        Thread errPump = pump(process.getErrorStream(), stderr, "finebi-stderr");

        long deadline = startedAt + timeoutSeconds * 1000L;
        boolean timedOut = false;
        while (true) {
            try {
                int code = process.exitValue();
                outPump.join(5000L);
                errPump.join(5000L);
                return new Result(code, stdout.toString(), stderr.toString(),
                        System.currentTimeMillis() - startedAt, false);
            } catch (IllegalThreadStateException stillRunning) {
                if (System.currentTimeMillis() > deadline) {
                    timedOut = true;
                    process.destroyForcibly();
                    break;
                }
                Thread.sleep(200L);
            }
        }
        outPump.join(3000L);
        errPump.join(3000L);
        return new Result(-1, stdout.toString(), stderr.toString(),
                System.currentTimeMillis() - startedAt, timedOut);
    }

    private static Thread pump(final InputStream in, final StringBuilder sink, String name) {
        Thread t = new Thread(new Runnable() {
            @Override
            public void run() {
                BufferedReader reader = null;
                try {
                    reader = new BufferedReader(new InputStreamReader(in, StandardCharsets.UTF_8));
                    String line;
                    while ((line = reader.readLine()) != null) {
                        sink.append(line).append('\n');
                    }
                } catch (IOException ignore) {
                    // 子进程被强杀时会走到这里，属于预期
                } finally {
                    if (reader != null) {
                        try {
                            reader.close();
                        } catch (IOException ignore) {
                        }
                    }
                }
            }
        }, name);
        t.setDaemon(true);
        t.start();
        return t;
    }

    /** 一轮执行的结果。 */
    public static final class Result {
        private final int exitCode;
        private final String stdout;
        private final String stderr;
        private final long millis;
        private final boolean timedOut;

        Result(int exitCode, String stdout, String stderr, long millis, boolean timedOut) {
            this.exitCode = exitCode;
            this.stdout = stdout;
            this.stderr = stderr;
            this.millis = millis;
            this.timedOut = timedOut;
        }

        public int getExitCode() {
            return exitCode;
        }

        public String getStdout() {
            return stdout;
        }

        public String getStderr() {
            return stderr;
        }

        public long getMillis() {
            return millis;
        }

        public boolean isTimedOut() {
            return timedOut;
        }

        /** 退出码 0 且 stdout 里 ok=true 才算真成功。 */
        public boolean isOk() {
            return exitCode == 0 && stdout.contains("\"ok\": true");
        }

        /** 粗提取产出文件路径（不引 JSON 库，够调度层打日志用）。 */
        public List<String> getFiles() {
            List<String> files = new ArrayList<String>();
            Matcher m = Pattern.compile("\"([^\"]*\\.xlsx)\"").matcher(stdout);
            while (m.find()) {
                files.add(m.group(1));
            }
            return files;
        }

        /** 失败时给调度平台的错误摘要。 */
        public String describe() {
            if (timedOut) {
                return "TIMEOUT after " + millis + "ms";
            }
            if (isOk()) {
                return "OK in " + millis + "ms, files=" + getFiles().size();
            }
            return "FAILED exit=" + exitCode + ", stderr tail=" + tail(stderr, 400);
        }

        private static String tail(String s, int n) {
            if (s == null) {
                return "";
            }
            return s.length() <= n ? s : s.substring(s.length() - n);
        }
    }
}
