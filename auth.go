package main

import (
	"context"
	"encoding/base64"
	"encoding/json"
	"fmt"
	"os"

	"golang.org/x/oauth2"
	"golang.org/x/oauth2/google"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/option"
	"google.golang.org/api/sheets/v4"
)

const (
	SheetsScope = "https://www.googleapis.com/auth/spreadsheets"
	DriveScope  = "https://www.googleapis.com/auth/drive"
)

var requiredScopes = []string{SheetsScope, DriveScope}

type AuthConfig struct {
	CredentialsConfig  string
	ServiceAccountPath string
	CredentialsPath    string
	TokenPath          string
	DriveFolderID      string
}

func LoadAuthConfig() *AuthConfig {
	return &AuthConfig{
		CredentialsConfig:  os.Getenv("CREDENTIALS_CONFIG"),
		ServiceAccountPath: os.Getenv("SERVICE_ACCOUNT_PATH"),
		CredentialsPath:    getEnvOrDefault("CREDENTIALS_PATH", "credentials.json"),
		TokenPath:          getEnvOrDefault("TOKEN_PATH", "token.json"),
		DriveFolderID:      os.Getenv("DRIVE_FOLDER_ID"),
	}
}

func getEnvOrDefault(key, defaultValue string) string {
	if value := os.Getenv(key); value != "" {
		return value
	}
	return defaultValue
}

func (ac *AuthConfig) GetCredentials(ctx context.Context) (*oauth2.Token, []byte, error) {
	// Priority 1: CREDENTIALS_CONFIG (Base64 encoded)
	if ac.CredentialsConfig != "" {
		fmt.Println("Using CREDENTIALS_CONFIG")
		credBytes, err := base64.StdEncoding.DecodeString(ac.CredentialsConfig)
		if err != nil {
			return nil, nil, fmt.Errorf("failed to decode CREDENTIALS_CONFIG: %w", err)
		}
		return nil, credBytes, nil
	}

	// Priority 2: SERVICE_ACCOUNT_PATH or GOOGLE_APPLICATION_CREDENTIALS
	serviceAcctPath := ac.ServiceAccountPath
	if serviceAcctPath == "" {
		serviceAcctPath = os.Getenv("GOOGLE_APPLICATION_CREDENTIALS")
	}

	if serviceAcctPath != "" && fileExists(serviceAcctPath) {
		fmt.Printf("Using service account: %s\n", serviceAcctPath)
		credBytes, err := os.ReadFile(serviceAcctPath)
		if err != nil {
			return nil, nil, fmt.Errorf("failed to read service account file: %w", err)
		}
		if ac.DriveFolderID != "" {
			fmt.Printf("Working with Google Drive folder ID: %s\n", ac.DriveFolderID)
		}
		return nil, credBytes, nil
	}

	// Priority 3: OAuth with CREDENTIALS_PATH
	if fileExists(ac.CredentialsPath) {
		fmt.Println("Using OAuth authentication flow")

		credBytes, err := os.ReadFile(ac.CredentialsPath)
		if err != nil {
			return nil, nil, fmt.Errorf("failed to read credentials file: %w", err)
		}

		config, err := google.ConfigFromJSON(credBytes, requiredScopes...)
		if err != nil {
			return nil, nil, fmt.Errorf("failed to parse credentials: %w", err)
		}

		token, err := ac.getTokenFromFile()
		if err != nil || !token.Valid() {
			token, err = ac.getTokenFromWeb(config)
			if err != nil {
				return nil, nil, fmt.Errorf("failed to get OAuth token: %w", err)
			}
			if err := ac.saveToken(token); err != nil {
				fmt.Printf("Warning: failed to save token: %v\n", err)
			}
		}

		return token, credBytes, nil
	}

	// Priority 4: Application Default Credentials
	fmt.Println("Attempting to use Application Default Credentials (ADC)")
	fmt.Println("ADC will check: GOOGLE_APPLICATION_CREDENTIALS, gcloud auth, and metadata service")

	creds, err := google.FindDefaultCredentials(ctx, requiredScopes...)
	if err != nil {
		return nil, nil, fmt.Errorf("all authentication methods failed: %w", err)
	}

	fmt.Println("Successfully authenticated using ADC")
	return nil, creds.JSON, nil
}

func (ac *AuthConfig) CreateServices(ctx context.Context) (*sheets.Service, *drive.Service, error) {
	token, credBytes, err := ac.GetCredentials(ctx)
	if err != nil {
		return nil, nil, err
	}

	var opts []option.ClientOption

	if token == nil && credBytes != nil {
		var credMap map[string]any
		if err := json.Unmarshal(credBytes, &credMap); err == nil {
			if credType, ok := credMap["type"].(string); ok && credType == "service_account" {
				creds, err := google.CredentialsFromJSON(ctx, credBytes, requiredScopes...)
				if err != nil {
					return nil, nil, fmt.Errorf("failed to create service account credentials: %w", err)
				}
				opts = append(opts, option.WithCredentials(creds))
			} else {
				creds, err := google.CredentialsFromJSON(ctx, credBytes, requiredScopes...)
				if err != nil {
					return nil, nil, fmt.Errorf("failed to create credentials: %w", err)
				}
				opts = append(opts, option.WithCredentials(creds))
			}
		} else {
			return nil, nil, fmt.Errorf("failed to parse credentials JSON: %w", err)
		}
	} else if token != nil && credBytes != nil {
		config, err := google.ConfigFromJSON(credBytes, requiredScopes...)
		if err != nil {
			return nil, nil, fmt.Errorf("failed to parse OAuth config: %w", err)
		}
		client := config.Client(ctx, token)
		opts = append(opts, option.WithHTTPClient(client))
	}

	sheetsService, err := sheets.NewService(ctx, opts...)
	if err != nil {
		return nil, nil, fmt.Errorf("failed to create sheets service: %w", err)
	}

	driveService, err := drive.NewService(ctx, opts...)
	if err != nil {
		return nil, nil, fmt.Errorf("failed to create drive service: %w", err)
	}

	return sheetsService, driveService, nil
}

func (ac *AuthConfig) getTokenFromFile() (*oauth2.Token, error) {
	f, err := os.Open(ac.TokenPath)
	if err != nil {
		return nil, err
	}
	defer f.Close()

	token := &oauth2.Token{}
	if err := json.NewDecoder(f).Decode(token); err != nil {
		return nil, err
	}
	return token, nil
}

func (ac *AuthConfig) getTokenFromWeb(config *oauth2.Config) (*oauth2.Token, error) {
	authURL := config.AuthCodeURL("state-token", oauth2.AccessTypeOffline)
	fmt.Printf("Go to the following link in your browser:\n%v\n", authURL)
	fmt.Print("Enter authorization code: ")

	var authCode string
	if _, err := fmt.Scan(&authCode); err != nil {
		return nil, fmt.Errorf("failed to read authorization code: %w", err)
	}

	token, err := config.Exchange(context.Background(), authCode)
	if err != nil {
		return nil, fmt.Errorf("failed to exchange authorization code: %w", err)
	}
	return token, nil
}

func (ac *AuthConfig) saveToken(token *oauth2.Token) error {
	f, err := os.OpenFile(ac.TokenPath, os.O_RDWR|os.O_CREATE|os.O_TRUNC, 0600)
	if err != nil {
		return err
	}
	defer f.Close()

	return json.NewEncoder(f).Encode(token)
}

func fileExists(path string) bool {
	_, err := os.Stat(path)
	return err == nil
}
