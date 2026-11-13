// BGI Underbilling Report launcher.
//
// A single EXE that serves the embedded report page to this computer only and
// opens it in an app window. It has no third-party dependencies and makes no
// outbound network connections.
//
// Security model
//   - Listens on 127.0.0.1 on a random port; never on a network interface.
//   - Every request must carry a random per-run token and the exact Host header,
//     so other websites and other local pages cannot read the data.
//   - Serves exactly two kinds of file from the folder the EXE sits in: the
//     newest UNDERBILLING*.xlsx export and the newest *Schedule*.csv. No paths.
//   - Strict Content-Security-Policy with a per-response nonce.
package main

import (
	"crypto/rand"
	"crypto/subtle"
	_ "embed"
	"encoding/base64"
	"encoding/json"
	"flag"
	"fmt"
	"io"
	"net"
	"net/http"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"sync"
	"time"
)

//go:embed ui/index.html
var indexHTML string

var version = "dev"

const (
	maxFileBytes = 60 << 20
	idleLimit    = 3 * time.Minute // background tabs throttle timers, so be generous
	byeGrace     = 4 * time.Second // lets a page reload re-attach before we exit
)

type fileInfo struct {
	Name     string `json:"name"`
	Size     int64  `json:"size"`
	Modified string `json:"modified"`
}

type server struct {
	token   string
	host    string
	dataDir string

	mu       sync.Mutex
	lastSeen time.Time
	byeAt    time.Time
}

func randomToken(n int) (string, error) {
	b := make([]byte, n)
	if _, err := rand.Read(b); err != nil {
		return "", err
	}
	return base64.RawURLEncoding.EncodeToString(b), nil
}

// newest returns the most recently modified regular file in dir that passes match.
func newest(dir string, match func(lower string) bool) *fileInfo {
	entries, err := os.ReadDir(dir)
	if err != nil {
		return nil
	}
	var found []fileInfo
	for _, e := range entries {
		name := e.Name()
		if strings.HasPrefix(name, "~$") || strings.HasPrefix(name, ".") || !match(strings.ToLower(name)) {
			continue
		}
		info, err := os.Lstat(filepath.Join(dir, name))
		if err != nil || !info.Mode().IsRegular() || info.Size() > maxFileBytes {
			continue
		}
		found = append(found, fileInfo{Name: name, Size: info.Size(), Modified: info.ModTime().UTC().Format(time.RFC3339)})
	}
	if len(found) == 0 {
		return nil
	}
	sort.Slice(found, func(i, j int) bool { return found[i].Modified > found[j].Modified })
	return &found[0]
}

func isExport(lower string) bool {
	return strings.HasSuffix(lower, ".xlsx") && strings.Contains(lower, "underbilling")
}

func isSchedule(lower string) bool {
	return strings.HasSuffix(lower, ".csv") && strings.Contains(lower, "schedule")
}

func (s *server) touch() {
	s.mu.Lock()
	s.lastSeen = time.Now()
	s.byeAt = time.Time{}
	s.mu.Unlock()
}

func (s *server) tokenOK(got string) bool {
	return subtle.ConstantTimeCompare([]byte(got), []byte(s.token)) == 1
}

// guard rejects anything that is not a same-origin request from our own page.
func (s *server) guard(next http.HandlerFunc, needHeaderToken bool) http.HandlerFunc {
	return func(w http.ResponseWriter, r *http.Request) {
		h := w.Header()
		h.Set("X-Content-Type-Options", "nosniff")
		h.Set("Referrer-Policy", "no-referrer")
		h.Set("Cache-Control", "no-store")
		h.Set("X-Frame-Options", "DENY")
		h.Set("Cross-Origin-Opener-Policy", "same-origin")
		h.Set("Cross-Origin-Resource-Policy", "same-origin")
		if r.Host != s.host {
			http.Error(w, "forbidden", http.StatusForbidden)
			return
		}
		if o := r.Header.Get("Origin"); o != "" && o != "http://"+s.host {
			http.Error(w, "forbidden", http.StatusForbidden)
			return
		}
		tok := r.URL.Query().Get("t")
		if needHeaderToken {
			tok = r.Header.Get("X-BGI-Token")
		}
		if !s.tokenOK(tok) {
			http.Error(w, "forbidden", http.StatusForbidden)
			return
		}
		next(w, r)
	}
}

func (s *server) handleIndex(w http.ResponseWriter, r *http.Request) {
	if r.URL.Path != "/" || r.Method != http.MethodGet {
		http.NotFound(w, r)
		return
	}
	nonce, err := randomToken(18)
	if err != nil {
		http.Error(w, "internal error", http.StatusInternalServerError)
		return
	}
	csp := "default-src 'none'; script-src 'nonce-" + nonce + "'; style-src 'nonce-" + nonce + "'; style-src-attr 'unsafe-inline'; " +
		"connect-src 'self'; img-src data:; base-uri 'none'; form-action 'none'; frame-ancestors 'none'"
	w.Header().Set("Content-Security-Policy", csp)
	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	s.touch()
	io.WriteString(w, strings.ReplaceAll(indexHTML, "__CSP_NONCE__", nonce))
}

func (s *server) handleFiles(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodGet {
		http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
		return
	}
	s.touch()
	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(map[string]*fileInfo{
		"export":   newest(s.dataDir, isExport),
		"schedule": newest(s.dataDir, isSchedule),
	})
}

func (s *server) handleFile(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodGet {
		http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
		return
	}
	s.touch()
	name := r.URL.Query().Get("name")
	// Only the two files the listing offers can be fetched; anything else is refused.
	var allowed *fileInfo
	for _, c := range []*fileInfo{newest(s.dataDir, isExport), newest(s.dataDir, isSchedule)} {
		if c != nil && c.Name == name {
			allowed = c
		}
	}
	if allowed == nil || name != filepath.Base(name) {
		http.Error(w, "not found", http.StatusNotFound)
		return
	}
	f, err := os.Open(filepath.Join(s.dataDir, name))
	if err != nil {
		http.Error(w, "not found", http.StatusNotFound)
		return
	}
	defer f.Close()
	w.Header().Set("Content-Type", "application/octet-stream")
	io.Copy(w, io.LimitReader(f, maxFileBytes))
}

func (s *server) handlePing(w http.ResponseWriter, r *http.Request) {
	s.touch()
	w.WriteHeader(http.StatusNoContent)
}

func (s *server) handleBye(w http.ResponseWriter, r *http.Request) {
	s.mu.Lock()
	s.byeAt = time.Now()
	s.mu.Unlock()
	w.WriteHeader(http.StatusNoContent)
}

// watch exits the process once the report window is gone.
func (s *server) watch() {
	for range time.Tick(time.Second) {
		s.mu.Lock()
		idle := time.Since(s.lastSeen)
		closed := !s.byeAt.IsZero() && time.Since(s.byeAt) > byeGrace
		s.mu.Unlock()
		if closed || idle > idleLimit {
			os.Exit(0)
		}
	}
}

func main() {
	dir := flag.String("dir", "", "folder to read the export and schedule from (default: the folder this program is in)")
	noOpen := flag.Bool("no-open", false, "do not open a window; print the address instead")
	showVersion := flag.Bool("version", false, "print the version and exit")
	flag.Parse()
	if *showVersion {
		fmt.Println(version)
		return
	}

	dataDir := *dir
	if dataDir == "" {
		exe, err := os.Executable()
		if err != nil {
			fatal("Could not work out where this program is running from: " + err.Error())
		}
		if resolved, err := filepath.EvalSymlinks(exe); err == nil {
			exe = resolved
		}
		dataDir = filepath.Dir(exe)
	}

	token, err := randomToken(32)
	if err != nil {
		fatal("Could not start securely: " + err.Error())
	}
	ln, err := net.Listen("tcp4", "127.0.0.1:0")
	if err != nil {
		fatal("Could not start the report on this computer: " + err.Error())
	}
	s := &server{token: token, host: ln.Addr().String(), dataDir: dataDir, lastSeen: time.Now()}

	mux := http.NewServeMux()
	mux.HandleFunc("/", s.guard(s.handleIndex, false))
	mux.HandleFunc("/api/files", s.guard(s.handleFiles, true))
	mux.HandleFunc("/api/file", s.guard(s.handleFile, true))
	mux.HandleFunc("/api/ping", s.guard(s.handlePing, true))
	mux.HandleFunc("/api/bye", s.guard(s.handleBye, false))

	url := "http://" + s.host + "/?t=" + token
	if *noOpen {
		fmt.Println(url)
	} else if err := openWindow(url); err != nil {
		fatal("The report is ready but a browser window could not be opened.\n\nOpen Microsoft Edge and go to:\n" + url)
	}
	go s.watch()

	srv := &http.Server{Handler: mux, ReadHeaderTimeout: 5 * time.Second, MaxHeaderBytes: 16 << 10}
	if err := srv.Serve(ln); err != nil {
		fatal("The report stopped unexpectedly: " + err.Error())
	}
}
